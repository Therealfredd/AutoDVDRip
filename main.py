import os
import subprocess
from time import sleep
import psutil
import ctypes
import argparse
import re
import glob


def get_volume_name(drive_path):
    volume_name_buffer = ctypes.create_unicode_buffer(1024)
    serial_number = None
    max_component_length = None
    file_system_flags = None
    file_system_name_buffer = ctypes.create_unicode_buffer(1024)

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
    return ""

def detect_dvd_drive():
    for drive in psutil.disk_partitions():
        if 'cdrom' in drive.opts.lower() and drive.device.upper() != 'C:\\':
            volume_name = get_volume_name(drive.device)
            if volume_name:  # Check if there's a valid volume name
                return drive.device
    return None


def get_volume_information(drive_path):
    volume_name_buffer = ctypes.create_unicode_buffer(1024)
    serial_number = ctypes.c_ulong()
    max_component_length = ctypes.c_ulong()
    file_system_flags = ctypes.c_ulong()
    file_system_name_buffer = ctypes.create_unicode_buffer(1024)

    if not drive_path.endswith("\\"):
        drive_path += "\\"

    success = ctypes.windll.kernel32.GetVolumeInformationW(
        ctypes.c_wchar_p(drive_path),
        volume_name_buffer,
        ctypes.sizeof(volume_name_buffer),
        ctypes.byref(serial_number),
        ctypes.byref(max_component_length),
        ctypes.byref(file_system_flags),
        file_system_name_buffer,
        ctypes.sizeof(file_system_name_buffer)
    )

    if success:
        return volume_name_buffer.value, serial_number.value
    else:
        return "", None

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


def convert_dvd(dvd_drive, output_folder, makemkv_path=r"C:\Program Files (x86)\MakeMKV\makemkvcon64.exe",
                handbrake_path=r"C:\Program Files\Handbrake\HandBrakeCLI.exe", minlength=900):

    info_cmd = f'"{makemkv_path}" info dev:{dvd_drive} --minlength={minlength} -r'
    result = subprocess.run(info_cmd, shell=True, text=True, capture_output=True)
    title_id = None
    dvd_name = None
    total_seconds_2 = 0

    name_pattern = re.compile(r'CINFO:2,0,"(.+?)"')
    title_pattern = re.compile(r'MSG:3028,\d+,\d+,"Title #(\d+) was added \(\d+ cell\(s\), (\d+):(\d+):(\d+)\)"')

    for line in result.stdout.splitlines():
        name_match = name_pattern.search(line)
        if name_match:
            dvd_name = name_match.group(1)

        title_match = title_pattern.search(line)
        if title_match:
            current_title_id = title_match.group(1)
            hours = int(title_match.group(2))
            minutes = int(title_match.group(3))
            seconds = int(title_match.group(4))
            total_seconds = hours * 3600 + minutes * 60 + seconds
            if total_seconds > total_seconds_2:
                total_seconds_2 = total_seconds
            if total_seconds >= minlength and not title_id:
                title_id = current_title_id
                print(f"Selected Title ID: {title_id} with duration {total_seconds} seconds")

    if not title_id:
        print("No suitable title found.")
        return

    if not dvd_name:
        print("DVD name not found, using default name 'Unknown_DVD'")
        dvd_name = "Unknown_DVD"

    dvd_name = sanitize_dvd_name(dvd_name)
    destination = os.path.join(output_folder, dvd_name)
    try:
        os.makedirs(destination, exist_ok=True)
    except Exception as e:
        print(f"Error creating directory: {e}")


    makemkv_cmd = f'"{makemkv_path}" mkv dev:{dvd_drive} all "{destination}" --minlength={total_seconds_2}'
    print(f"Running MakeMKV: {makemkv_cmd}")
    subprocess.run(makemkv_cmd, shell=True)

    process_mkvs(destination, dvd_name, handbrake_path)


def process_mkvs(destination, dvd_name, handbrake_path):
    mkv_files = glob.glob(os.path.join(destination, '*.mkv'))
    for i, mkv_file in enumerate(mkv_files, start=0):
        mp4_file = os.path.join(destination, f"{dvd_name}_{i}.mp4")
        handbrake_cmd = f'"{handbrake_path}" -i "{mkv_file}" -o "{mp4_file}" --preset="H.265 NVENC 1080p"'
        print(f"Running HandBrakeCLI: {handbrake_cmd}")
        subprocess.run(handbrake_cmd, shell=True)
        os.remove(mkv_file)
        print(f"Deleted original MKV file: {mkv_file}")


def sanitize_dvd_name(name):
    illegal_chars = '<>:"/\\|?*'
    for char in illegal_chars:
        name = name.replace(char, '')
    return name

def main(output_folder,watch_drive):
    processed_dvds = set()

    try:
        with open("processed_dvds.log", "r") as file:
            processed_dvds.update(file.read().splitlines())
    except FileNotFoundError:
        pass

    while True:
        if not watch_drive:
            dvd_drive = detect_dvd_drive()
        else:
            dvd_drive = args.watch_drive + "\\"
        if dvd_drive:
            _, serial_number = get_volume_information(dvd_drive)
            if str(serial_number) in processed_dvds:
                print(f"DVD with serial {serial_number} has already been processed. Skipping...")
                sleep(5)
                continue
            convert_dvd(dvd_drive, output_folder)

            if serial_number:
                processed_dvds.add(str(serial_number))
                with open("processed_dvds.log", "a") as file:
                    file.write(f"{serial_number}\n")
            else:
                print("Unable to retrieve DVD serial number, processing may be duplicated in future runs.")

        else:
            print("No DVD detected. Checking again in 15 seconds...")
        sleep(15)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Automate DVD conversion to MP4.')
    parser.add_argument('output_folder', nargs='?', default='output', type=str,
                        help='The path to the output folder where the DVD will be converted and saved.')
    parser.add_argument('watch_drive', type=str, default='',
                        help='The drive letter or path to watch for DVD insertion. Default=automatic detection.')
    args = parser.parse_args()
    main(args.output_folder, args.watch_drive)
