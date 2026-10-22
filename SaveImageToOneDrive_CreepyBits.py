import numpy as np
from PIL import Image
from PIL.PngImagePlugin import PngInfo
import requests
import os
import folder_paths
import json
import threading
from urllib.parse import quote

import msal

NODE_DIR = os.path.dirname(os.path.abspath(__file__))

# Your Azure app's "Application (client) ID". Set it in onedrive_config.json next to this file
# ({"client_id": "..."}), or with the ONEDRIVE_CLIENT_ID environment variable, or paste it here.
CLIENT_ID = ""
CONFIG_FILE = os.path.join(NODE_DIR, "onedrive_config.json")
AUTHORITY = "https://login.microsoftonline.com/consumers"
SCOPES = ["https://graph.microsoft.com/Files.ReadWrite", "https://graph.microsoft.com/User.Read"]

TOKEN_CACHE_FILE = os.path.join(NODE_DIR, "onedrive_msal_token_cache.bin")

UPLOAD_TIMEOUT_SECONDS = 300

_msal_app = None
_msal_lock = threading.Lock()


def get_client_id():
    client_id = os.environ.get("ONEDRIVE_CLIENT_ID", "").strip() or CLIENT_ID.strip()
    if not client_id and os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                client_id = str(json.load(f).get("client_id", "")).strip()
        except Exception as e:
            print(f"[OneDrive Auth] Could not read {CONFIG_FILE}: {e}")
    return client_id


def get_msal_app():
    """Created on first use rather than at import, so a missing client ID doesn't stop ComfyUI from loading the node.
    The token cache is loaded from disk once and then kept in memory."""
    global _msal_app
    if _msal_app is None:
        client_id = get_client_id()
        if not client_id:
            return None
        _msal_app = msal.PublicClientApplication(client_id, authority=AUTHORITY, token_cache=load_msal_token_cache())
    return _msal_app


def load_msal_token_cache():
    cache = msal.SerializableTokenCache()
    if os.path.exists(TOKEN_CACHE_FILE):
        try:
            with open(TOKEN_CACHE_FILE, "r") as f:
                cache.deserialize(f.read())
        except Exception as e:
            print(f"[OneDrive Auth] Error loading MSAL token cache: {e}. Starting with an empty cache.")
            try:
                os.remove(TOKEN_CACHE_FILE)
            except OSError as oe:
                print(f"[OneDrive Auth] Error removing corrupted token cache file: {oe}")
    return cache

def save_msal_token_cache(cache):
    if cache.has_state_changed:
        try:
            with open(TOKEN_CACHE_FILE, "w") as f:
                f.write(cache.serialize())
            cache.has_state_changed = False
        except Exception as e:
            print(f"[OneDrive Auth] Error saving MSAL token cache: {e}")

def get_onedrive_access_token_msal():
    msal_app = get_msal_app()
    if msal_app is None:
        message = ("OneDrive client ID is not set. Put {\"client_id\": \"<your Azure app client ID>\"} in "
                   f"{CONFIG_FILE}, or set the ONEDRIVE_CLIENT_ID environment variable.")
        print(f"[OneDrive Auth] {message}")
        return None, message

    with _msal_lock:
        accounts = msal_app.get_accounts()
        result = None
        status_message = None

        if accounts:
            result = msal_app.acquire_token_silent(SCOPES, account=accounts[0])

        if not result:
            flow = msal_app.initiate_device_flow(scopes=SCOPES)

            if "user_code" not in flow:
                error_message = f"Failed to create device flow. Error: {flow.get('error_description', str(flow))}"
                print(f"[OneDrive Auth] {error_message}")
                return None, error_message

            device_flow_instructions = (
                f"ONEDRIVE AUTH NEEDED (Device Flow):\n"
                f"1. Open browser to: {flow['verification_uri']}\n"
                f"2. Enter code: {flow['user_code']}\n"
                f"3. Sign in & grant permissions.\n"
                f"The node waits here until you have signed in, then continues the upload."
            )
            print("----------------------------------------------------------------------")
            print(device_flow_instructions)
            print("----------------------------------------------------------------------")
            status_message = device_flow_instructions.replace("\n", " - ")

            try:
                result = msal_app.acquire_token_by_device_flow(flow)
            except Exception as e_device_flow:
                print(f"[OneDrive Auth] Error during acquire_token_by_device_flow: {e_device_flow}")
                status_message = f"Error during device flow: {e_device_flow}"
                result = None

        if result and "access_token" in result:
            save_msal_token_cache(msal_app.token_cache)
            return result["access_token"], "OneDrive authentication successful."
        elif result and "error_description" in result:
            error_desc = f"Failed to acquire token. Error: {result.get('error')}, Desc: {result.get('error_description')}"
            print(f"[OneDrive Auth] {error_desc}")
            return None, error_desc
        else:
            if status_message:
                return None, status_message
            unknown_error_msg = "[OneDrive Auth] Could not acquire token. Unknown error or device flow not completed."
            print(unknown_error_msg)
            return None, unknown_error_msg

class SaveImageToOneDrive_CreepyBits:
    # Per-prefix counters. OneDrive renames on conflict, so these only need to be unique
    # within a session (the temp-folder files they used to be tracked with were wiped on restart anyway).
    _counters = {}
    _counter_lock = threading.Lock()

    def __init__(self):
        self.output_dir = folder_paths.get_temp_directory()
        self.type = "temp"

    @classmethod
    def INPUT_TYPES(s):
        return {
            "required": {
                "images": ("IMAGE", ),
                "filename_prefix": ("STRING", {"default": "ComfyUI"}),
                "onedrive_folder_path": ("STRING", {"default": "ComfyUI_Uploads"}),
                "file_format": (["PNG", "JPEG", "WEBP"], {"default": "PNG"}),
                "quality": ("INT", {"default": 90, "min": 1, "max": 100, "step": 1}),
            },
            "hidden": {
                "prompt": "PROMPT",
                "extra_pnginfo": "EXTRA_PNGINFO"
            },
        }

    RETURN_TYPES = ("STRING",)
    RETURN_NAMES = ("status_message",)
    OUTPUT_NODE = True
    FUNCTION = "execute"
    CATEGORY = "Creepybits/Cloud"

    def tensor_to_pil(self, tensor_image):
        if tensor_image.ndim == 4:
            tensor_image = tensor_image[0]
        img_np = tensor_image.cpu().numpy() * 255.0
        pil_img = Image.fromarray(np.clip(img_np, 0, 255).astype(np.uint8))
        return pil_img

    def get_next_indexed_filename(self, filename_prefix, file_ext):
        key = (filename_prefix.lower(), file_ext.lower())
        with self._counter_lock:
            next_index = self._counters.get(key, 0) + 1
            self._counters[key] = next_index
        return f"{filename_prefix}_{next_index:05d}.{file_ext}"

    def execute(self, images, onedrive_folder_path, filename_prefix, file_format, quality, prompt=None, extra_pnginfo=None):
        access_token, auth_status_message = get_onedrive_access_token_msal()

        if not access_token:
            print(f"[SaveToOneDrive] Authentication failed or requires user action. Message: {auth_status_message}")
            display_msg = auth_status_message.replace("\n", " - ") if auth_status_message else \
                          "AUTH_ERROR: OneDrive authentication failed. Check console for details and re-queue prompt."
            return {"ui": {"string": [display_msg]}, "result": (display_msg,)}

        cleaned_folder_path = onedrive_folder_path.strip().strip('/')

        chosen_format_upper = file_format.upper()
        file_extension = "png"

        if chosen_format_upper == "JPEG":
            file_extension = "jpeg"
        elif chosen_format_upper == "WEBP":
            file_extension = "webp"

        mimetype_map = {
            "PNG": "image/png",
            "JPEG": "image/jpeg",
            "WEBP": "image/webp"
        }
        upload_mimetype = mimetype_map.get(chosen_format_upper, 'application/octet-stream')

        # PNG metadata is the same for every image in the batch, so build it once.
        metadata = None
        if chosen_format_upper == "PNG":
            metadata = PngInfo()
            if prompt is not None:
                metadata.add_text("prompt", json.dumps(prompt))
            if extra_pnginfo is not None:
                for k, v in extra_pnginfo.items():
                    metadata.add_text(k, json.dumps(v))

        all_uploaded_files = []
        all_previews = []

        os.makedirs(self.output_dir, exist_ok=True)

        # One connection for the whole batch instead of a new TLS handshake per image.
        with requests.Session() as session:
            session.headers.update({"Authorization": f"Bearer {access_token}", "Content-Type": upload_mimetype})

            for i, image_tensor in enumerate(images):
                indexed_filename_only = self.get_next_indexed_filename(filename_prefix, file_extension)
                temp_filepath_local = os.path.join(self.output_dir, indexed_filename_only)

                try:
                    pil_image_for_save = self.tensor_to_pil(image_tensor)

                    if chosen_format_upper == "JPEG" and pil_image_for_save.mode != "RGB":
                        pil_image_for_save = pil_image_for_save.convert("RGB")

                    if chosen_format_upper == "PNG":
                        pil_image_for_save.save(temp_filepath_local, pnginfo=metadata, compress_level=4)
                    else:
                        pil_image_for_save.save(temp_filepath_local, quality=quality)

                    with open(temp_filepath_local, 'rb') as f:
                        file_bytes = f.read()

                except Exception as e:
                    error_msg = f"ERROR processing image {i+1} for local save: {str(e)}"
                    print(f"[SaveToOneDrive] {error_msg}")
                    all_uploaded_files.append(f"Image {i+1} Failed: {error_msg}")
                    continue

                upload_item_path = f"{cleaned_folder_path}/{indexed_filename_only}" if cleaned_folder_path else indexed_filename_only

                # Encode the path so folder names with spaces, '#', '%' etc. upload to the right place.
                upload_url = f"https://graph.microsoft.com/v1.0/me/drive/root:/{quote(upload_item_path)}:/content?@microsoft.graph.conflictBehavior=rename"

                try:
                    response = session.put(upload_url, data=file_bytes, timeout=UPLOAD_TIMEOUT_SECONDS)
                    response.raise_for_status()

                    response_json = response.json()
                    file_web_url = response_json.get("webUrl", "N/A")
                    upload_status = f"SUCCESS: Uploaded '{response_json.get('name', indexed_filename_only)}'. URL: {file_web_url}"
                    all_uploaded_files.append(upload_status)

                    all_previews.append({"filename": indexed_filename_only, "subfolder": "", "type": "temp"})

                except requests.exceptions.HTTPError as http_err:
                    error_details = str(http_err.response.text if hasattr(http_err.response, 'text') and http_err.response.text else http_err)
                    try:
                        error_details = http_err.response.json().get("error", {}).get("message", error_details)
                    except Exception:
                        pass
                    upload_status = f"ERROR: HTTP {http_err.response.status_code} uploading '{indexed_filename_only}': {error_details}"
                    all_uploaded_files.append(upload_status)
                except Exception as e:
                    upload_status = f"ERROR: Unexpected error during upload of '{indexed_filename_only}': {str(e)}"
                    all_uploaded_files.append(upload_status)

                print(f"[SaveToOneDrive] {upload_status}")

        if all_uploaded_files:
            overall_status_message = " | ".join(all_uploaded_files)
        else:
            overall_status_message = "No images were uploaded due to errors."

        ui_return_data = {}
        if all_previews:
            ui_return_data["images"] = all_previews
        if overall_status_message:
            ui_return_data["string"] = [overall_status_message]

        return {"ui": ui_return_data, "result": (overall_status_message,)}

NODE_CLASS_MAPPINGS = { "SaveImageToOneDrive_CreepyBits": SaveImageToOneDrive_CreepyBits }
NODE_DISPLAY_NAME_MAPPINGS = { "SaveImageToOneDrive_CreepyBits": "Save Image to OneDrive (Creepybits)" }
