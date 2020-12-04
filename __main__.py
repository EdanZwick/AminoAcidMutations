import os
import dataHandeling
from datetime import datetime

supported_extensions = {'.xls', '.xlsx', '.csv'}


def dispatch(target_path):
    if os.path.isdir(target_path):
        print('Analyzing all files in directory')
        frames = work_on_folder(target_path)
    else:
        print('Analyzing single file')
        frames = work_on_file(target_path)
        target_path = os.path.dirname(target_path)
    now = datetime.now().strftime('%D_%H%M%S').replace("/", "")
    file_name = 'RESULT_' + now + '.xlsx'
    output_path = os.path.join(target_path, file_name)
    dataHandeling.finalize(frames, output_path)
    print('\n\n\n\n')
    print('wrote output to file:', output_path)


def work_on_folder(target_path):
    frames = []
    for fileName in os.listdir(target_path):
        if fileName[0] == '~' or fileName[0] == '$':
            print('file', full_path, 'is a tmp file and is skipped', sep='')
            continue
        full_path = os.path.join(target_path, fileName)
        frames += work_on_file(full_path)
    return frames


def work_on_file(full_path):
    if not os.path.exists(full_path):
        print('File', full_path, 'does not exist', sep=' ')
        return []
    if not os.path.splitext(full_path)[1] in supported_extensions:
        print('File', full_path, 'has an unsupported extension', sep=' ')
        return []
    print('working on file', full_path, sep=' ')
    result = dataHandeling.analyze_report(full_path)
    if result is None:
        print('failed on path:', full_path)
        return []
    return [(os.path.splitext(os.path.basename(full_path))[0], result)]


def main():
    target_path = input('Please enter path to excel file or to folder:\n')
    if not os.path.exists(target_path):
        print('Invalid path, please fix')
        return
    dispatch(target_path)
    print('done')
    input("Press Enter to continue...")


if __name__ == '__main__':
    main()
