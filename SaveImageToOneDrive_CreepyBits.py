import torch
import numpy as np
from PIL import Image
from PIL.PngImagePlugin import PngInfo
import io
import requests
import os
import folder_paths
import json
import re
import time

import msal

CLIENT_ID = YOUR_CLIENt_ID
AUTHORITY = "https://login.microsoftonline.com/consumers"
SCOPES = ["https://graph.microsoft.com/Files.ReadWrite", "https://graph.microsoft.com/User.Read"]

TOKEN_CACHE_FILE = os.path.join(os.path.dirname(__file__), "onedrive_msal_token_cache.bin")


msal_app = msal.PublicClientApplication(
    CLIENT_ID,
    authority=AUTHORITY
)

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
    current_token_cache = load_msal_token_cache()
    msal_app.token_cache = current_token_cache

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
            status_message = error_message
            return None, status_message

        device_flow_instructions = (
            f"ONEDRIVE AUTH NEEDED (Device Flow):\n"
            f"1. Open browser to: {flow['verification_uri']}\n"
            f"2. Enter code: {flow['user_code']}\n"
            f"3. Sign in & grant permissions.\n"
            f"Node will then attempt to get token. Re-queue prompt after authorizing."
        )
        print("----------------------------------------------------------------------")
        print(device_flow_instructions)
        print("----------------------------------------------------------------------")
        status_message = device_flow_instructions.replace("\n", " - ") + " - Check console for details and re-queue."

        try:
            result = msal_app.acquire_token_by_device_flow(flow)
        except Exception as e_device_flow:
            print(f"[OneDrive Auth] Error during acquire_token_by_device_flow: {e_device_flow}")
            status_message = f"Error during device flow: {e_device_flow}"
            result = None

    if result and "access_token" in result:
        save_msal_token_cache(msal_app.token_cache)
        status_message = "OneDrive authentication successful."
        return result["access_token"], status_message
    elif result and "error_description" in result:
        error_desc = f"Failed to acquire token. Error: {result.get('error')}, Desc: {result.get('error_description')}"
        print(f"[OneDrive Auth] {error_desc}")
        status_message = error_desc
        return None, status_message
    else:
        if status_message:
            return None, status_message
        else:
            unknown_error_msg = "[OneDrive Auth] Could not acquire token. Unknown error or device flow not completed."
            print(unknown_error_msg)
            return None, unknown_error_msg

class SaveImageToOneDrive_CreepyBits:
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

    def get_next_indexed_filename(self, output_dir, filename_prefix, file_ext):
        index_tracker_dir = os.path.join(output_dir, "onedrive_node_index_tracker")
        os.makedirs(index_tracker_dir, exist_ok=True)

        pattern = re.compile(rf"^{re.escape(filename_prefix)}_(\d+)\.{re.escape(file_ext)}$", re.IGNORECASE)

        max_index = 0
        try:
            for f in os.listdir(index_tracker_dir):
                match = pattern.match(f)
                if match:
                    current_index = int(match.group(1))
                    if current_index > max_index:
                        max_index = current_index
        except FileNotFoundError:
            pass

        next_index = max_index + 1
        indexed_filename_only = f"{filename_prefix}_{next_index:05d}.{file_ext}"

        try:
            with open(os.path.join(index_tracker_dir, indexed_filename_only), 'wb') as f:
                f.write(b'')
        except Exception as e:
            print(f"[SaveToOneDrive] Warning: Could not create dummy index file {indexed_filename_only} in {index_tracker_dir}: {e}")

        return indexed_filename_only

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


        all_uploaded_files = []
        all_previews = []
        overall_status_message = ""

        temp_actual_image_dir = folder_paths.get_temp_directory()

        for i, image_tensor in enumerate(images):
            indexed_filename_only = self.get_next_indexed_filename(
                self.output_dir, filename_prefix, file_extension
            )
            temp_filepath_local = os.path.join(temp_actual_image_dir, indexed_filename_only)

            pil_image_for_save = None
            try:
                pil_image_for_save = self.tensor_to_pil(image_tensor)

                if chosen_format_upper == "JPEG":
                    if pil_image_for_save.mode == 'RGBA':
                        pil_image_for_save = pil_image_for_save.convert("RGB")
                elif chosen_format_upper == "WEBP":
                    if pil_image_for_save.mode == 'P':
                        pil_image_for_save = pil_image_for_save.convert("RGBA")
                    elif pil_image_for_save.mode == 'RGBA' and not pil_image_for_save.getchannel('A').getbbox():
                        pil_image_for_save = pil_image_for_save.convert("RGB")

                metadata = None
                if chosen_format_upper == "PNG":
                    metadata = PngInfo()
                    if prompt is not None:
                        metadata.add_text("prompt", json.dumps(prompt))
                    if extra_pnginfo is not None:
                        for k, v in extra_pnginfo.items():
                            metadata.add_text(k, json.dumps(v))

                os.makedirs(temp_actual_image_dir, exist_ok=True)
                if chosen_format_upper == "PNG":
                    pil_image_for_save.save(temp_filepath_local, pnginfo=metadata, compress_level=4)
                elif chosen_format_upper in ["JPEG", "WEBP"]:
                    pil_image_for_save.save(temp_filepath_local, quality=quality)
                else:
                    pil_image_for_save.save(temp_filepath_local)

                with open(temp_filepath_local, 'rb') as f:
                    file_bytes = f.read()

            except Exception as e:
                error_msg = f"ERROR processing image {i+1} for local save: {str(e)}"
                print(f"[SaveToOneDrive] {error_msg}")
                overall_status_message += f"Image {i+1} Failed: {error_msg}. "
                continue

            upload_item_path = f"{cleaned_folder_path}/{indexed_filename_only}" if cleaned_folder_path else indexed_filename_only

            upload_url = f"https://graph.microsoft.com/v1.0/me/drive/root:/{upload_item_path}:/content?@microsoft.graph.conflictBehavior=rename"

            headers = {
                "Authorization": f"Bearer {access_token}",
                "Content-Type": upload_mimetype
            }

            try:
                response = requests.put(upload_url, headers=headers, data=file_bytes)
                response.raise_for_status()

                response_json = response.json()
                file_web_url = response_json.get("webUrl", "N/A")
                upload_status = f"SUCCESS: Uploaded '{indexed_filename_only}'. URL: {file_web_url}"
                all_uploaded_files.append(upload_status)

                all_previews.append({"filename": indexed_filename_only, "subfolder": "", "type": "temp"})

            except requests.exceptions.HTTPError as http_err:
                error_details = str(http_err.response.text if hasattr(http_err.response, 'text') and http_err.response.text else http_err)
                try:
                    error_details = http_err.response.json().get("error", {}).get("message", error_details)
                except: pass
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
