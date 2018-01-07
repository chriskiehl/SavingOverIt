"""
A tool for cheating in the game Getting Over It
"""
import datetime
import itertools
import json
import os
import pickle
import sys
import time
import winreg
from bs4 import BeautifulSoup
from win32 import win32api, win32security

from gooey import Gooey, GooeyParser


priviliges_error = (
    'Looks like we can\'t access the registry where Getting Over It keeps its save data.'
    '\nMake sure you\'re running as an Administrator! Right click the SaveIt Icon and '
    'select "Run As Administrator"'
)

directory_error = (
    'Unable to create a folder for your save files!\n'
    'Make sure to run SaveOverIt! in a directory '
    'where you have full permissions'
)


@Gooey(
    program_name="Getting Over It!",
    poll_external_updates=True)
def main():
    parser = GooeyParser(description='Getting Over it without the "Hurt"')
    g = parser.add_argument_group()
    stuff = g.add_mutually_exclusive_group(
        required=True,
        gooey_options={
            'initial_selection': 0
        }
    )
    stuff.add_argument(
        '--save',
        metavar='Save Progress',
        action='store_true',
        help='Take a snap shot of your current progress!'
    )
    stuff.add_argument(
        '--load',
        metavar='Load Previous Save',
        help='Load a Previous save file',
        dest='filename',
        widget='Dropdown',
        choices=list_savefiles(),
        gooey_options={
            'validator': {
                'test': 'user_input != "Select Option"',
                'message': 'Choose a save file from the list'
            }
        }
    )

    args = parser.parse_args()

    if args.save:
        save_registry(collect_regvals(load_regkey()))
    else:
        replace_reg_values(load_regkey(), load_registry(os.path.join('saves', args.filename)))
        print('Successfully Loaded snapshot!')
        print('Make sure to fix your mouse sensitivity! ^_^\n')



def adjust_privileges():
    priv_flags = win32security.TOKEN_ADJUST_PRIVILEGES | win32security.TOKEN_QUERY
    token = win32security.OpenProcessToken(win32api.GetCurrentProcess(), priv_flags)
    privilege_id = win32security.LookupPrivilegeValue(None, "SeBackupPrivilege")
    win32security.AdjustTokenPrivileges(token, 0, [(privilege_id, win32security.SE_PRIVILEGE_ENABLED)])


def check_privileges():
    """
    See if we're running as an administrator by attempting
    to write something to the registry
    """
    try:
        winreg.SetValueEx(load_regkey(), 'com_gooey__tst_write', None, 4,0)
    except:
        show_error_modal(priviliges_error)
        sys.exit(1)


def load_regkey():
    return winreg.OpenKey(
        winreg.HKEY_CURRENT_USER,
        r"SOFTWARE\Bennett Foddy\Getting Over It",
        0,
        winreg.KEY_READ | winreg.KEY_SET_VALUE
    )


def replace_reg_values(key, data):
    for subkey, value, _type in data:
        winreg.SetValueEx(key, subkey, 0, _type, value)


def list_savefiles():
    return list(sorted(os.listdir('saves'), reverse=True))


def show_error_modal(error_msg):
    # wx imported locally so as not to interfere with Gooey
    import wx
    app = wx.App()
    dlg = wx.MessageDialog(None, error_msg, 'Error', wx.ICON_ERROR)
    dlg.ShowModal()
    dlg.Destroy()


def start_backup_process(duration):
    save_registry(collect_regvals(load_regkey()))
    print('backed up saved at {}'.format(datetime.datetime.now().isoformat()))
    print()
    time.sleep(duration)


def load_registry(filename):
    with open(filename,'rb') as f:
        return pickle.load(f)

def parse_int(s):
    return int(float(s))


def save_registry(data):
    next_number = '{:04d}'.format(len(list_savefiles()) + 1)
    now = datetime.datetime.now().strftime('%b %d, %Y - %H.%M%p')
    xpos, ypos = map(parse_int, get_pos_info(data))
    filename = 'Snapshot {} @ {} [-x-position {} -y-position {}].save'.format(next_number, now, xpos, ypos)
    with open(os.path.join('saves', filename), 'wb') as f:
        f.write(pickle.dumps(data))
    print('Saved {}!'.format(filename))


def get_pos_info(data):
    for subkey, value, _type in data:
        if subkey.startswith('SaveGame'):
            soup = BeautifulSoup(value.decode('utf-8'), 'html.parser').find('campos')
            return [soup.x.text, soup.y.text]


def collect_regvals(regkey):
    values = []
    for index in itertools.count(0):
        try:
            values.append(winreg.EnumValue(regkey, index))
        except OSError:
            break
    return values


def mk_savedir():
    try:
        os.mkdir('saves')
    except IOError as e:
        if not e.winerror == 183:  # already exists
            show_error_modal(directory_error)
            sys.exit(1)



if __name__ == '__main__':
    mk_savedir()
    adjust_privileges()
    check_privileges()
    if 'gooey-seed-ui' in sys.argv:
        print(json.dumps({'--load': list_savefiles()}))
    else:
        main()
