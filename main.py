import os
import subprocess
from time import sleep
import psutil
import ctypes
import argparse


def detect_dvd_drive():
    """Check all drives and return the drive letter of the DVD if present."""
    for drive in psutil.disk_partitions():
        if 'cdrom' in drive.opts.lower():
            return drive.device
    return None


def get_dvd_title(drive_path):
    volume_name_buffer = ctypes.create_unicode_buffer(1024)
    file_system_name_buffer = ctypes.create_unicode_buffer(1024)
    serial_number = None
    max_component_length = None
    file_system_flags = None

    if not drive_path.endswith("\\"):
        drive_path += "\\"

    success = ctypes.windll.kernel32.GetVolumeInformationW(
        ctypes.c_wchar_p(drive_path),
        volume_name_buffer,
        ctypes.sizeof(volume_name_buffer),
        serial_number,
        max_component_length,
        file_system_flags,
        file_system_name_buffer,
        ctypes.sizeof(file_system_name_buffer)
    )

    if success:
        return volume_name_buffer.value
    else:
        return ""


def convert_dvd(dvd_drive, output_folder, makemkv_path = r"C:\Program Files (x86)\MakeMKV\makemkvcon64.exe"):
    """Convert the DVD using MakeMKV and save it to the specified output folder."""
    title = get_dvd_title(dvd_drive)
    destination = os.path.join(output_folder, title)
    os.makedirs(destination, exist_ok=True)

    # Command to run MakeMKV, using the full path to the executable
    command = f'"{makemkv_path}" mkv dev:{dvd_drive} all "{destination}" --minlength=900'
    subprocess.run(command, shell=True)


def main(output_folder):
    while True:
        dvd_drive = detect_dvd_drive()
        if dvd_drive:
            title = get_dvd_title(dvd_drive)
            output_path = os.path.join(output_folder, title)

            if not os.path.exists(output_path):
                print(f"DVD detected in drive {dvd_drive} with title '{title}'. Converting...")
                convert_dvd(dvd_drive, output_folder)
            else:
                print(f"Folder for '{title}' already exists. Skipping conversion.")
        else:
            print("No DVD detected. Checking again in 30 seconds...")
        sleep(15)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Process some integers.')
    parser.add_argument('output_folder', type=str, help='The path to the output folder where the DVD will be converted and saved.')
    args = parser.parse_args()
    main(args.output_folder)
