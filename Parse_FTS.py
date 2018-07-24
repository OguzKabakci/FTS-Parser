# encoding: utf-8
#
# FTS_parse V1.0
# Date   : 02.07.2018
# Author : Oguz Kabakci

"""
FTS_parse document

1-) Get FTS document
2-) Take the variables
3-) Create a xml file with 10ms raster.
4-) Get ENV gile
5-) Get A2L file
6-) Give Read Write access to variables
"""

from tkinter import *
from tkinter import filedialog
from tkinter import messagebox

import docx2txt
import regex
import xml.etree.ElementTree as ElT
import os


class UserCancelled(Exception):
    pass


class FTS:
    @staticmethod
    def reg(doc):
        """
        Regex command to get 'variables, APV, CPV, APM, APT, NVV, NVM, NVT'
        :param doc: FTS document
        :return: member list
        """
        try:
            my_text = docx2txt.process(doc)
        except FileNotFoundError:
            sys.exit()
        my_text_eliminated = regex.findall('(?:\w+[a-z]_\w+)|(?:\w+(?:APV|CPV|APM|APT|NVV|NVM|NVT))', my_text)
        return my_text_eliminated

    @staticmethod
    def remove_duplicates(in_list: list):
        """
        Remove duplicates
        :param in_list: List of all variables
        :return: List without duplicate entries
        """
        return list(set(in_list))

    @staticmethod
    def write_file(member_list, complete_name):
        """
        Writes list in xml format
        :param member_list: member list
        :param complete_name: xml file path
        :return:
        """
        the_file = open(complete_name, 'w')
        # Create new xml properties
        the_file.write("<?xml version=\"1.0\" encoding=\"utf-8\"?>\n"
                       "<Screen VisuVersion=\"\" UserComment=\"\" DefRasterName=\"\" UtcDate=\"\">\n"
                       "  <Elements>\n")
        # Add the variables with 10ms raster
        for item in member_list:
            the_file.write("    <Element Raster=\"10ms\" Period=\"10\" Name=\"%s\"/>\n" % item)
        the_file.write("  </Elements>\n"
                       "</Screen>")

    @staticmethod
    def get_open_file_name(file_title, file_types):
        root = Tk()
        # Get rid of unnecessary pop up window
        root.withdraw()
        # Open a browser to select file
        root.filename = filedialog.askopenfilename(title=file_title,
                                                   filetypes=(file_types, ("All files", "*.*")))

        if root.filename == "":
            raise UserCancelled("User cancelled the pop up.")
        else:
            return root.filename

    @staticmethod
    def select_file(file_title, file_types):
        """
        Selecting the path of FTS file
        :return: FTS file path
        """
        # Get rid of unnecessary pop up window
        Tk().withdraw()
        try:
            file_name = screen.get_open_file_name(file_title, file_types)
        except UserCancelled as e_message:
            print(e_message)
            sys.exit()
        return file_name

    @staticmethod
    def get_save_file_name(file_path):
        root = Tk()
        # Get rid of unnecessary pop up window
        root.withdraw()

        # Open a browser to save file
        root.filename = filedialog.asksaveasfilename(initialdir=file_path, title="Step 2 of 4     Save XML file",
                                                     filetypes=(("xml files", "*.xml"), ("all files", "*.*")))

        if root.filename == "":
            raise UserCancelled("User cancelled the pop up.")
        else:
            return root.filename

    @staticmethod
    def save_xml(file_path):
        """
        Selecting the path of xml file
        :param file_path: FTS file path
        :return: xml file path with extension
        """
        try:
            file_name = screen.get_save_file_name(file_path)
        except UserCancelled as e_message:
            print(e_message)
            sys.exit()
        # Check the file name if it does not have '.xml' part, add it
        if '.xml' in file_name:
            complete_name = file_name
        else:
            complete_name = file_name + '.xml'
        return complete_name

    @staticmethod
    def check_a2l(env_path_temp, a2l_path_temp, list_temp):
        """
        check a2l file
        :param env_path_temp: Full path of env file
        :param a2l_path_temp: Full path of a2l file
        :param list_temp: The member list
        :return:
        """
        a2l_is_found = False
        # Get XML as tree
        tree = ElT.parse(env_path_temp)
        root = tree.getroot()
        # Find 'content' element as a sub tree
        root_sub_element_item = root.find("./content/items/content")
        # Find 'item' sub elements under 'content'
        for item_sub_element in root_sub_element_item:
            # Find specific 'item' with respect to a2l path
            if item_sub_element.find('file').text == a2l_path_temp:
                screen.modify_a2l_part(item_sub_element, list_temp)
                a2l_is_found = True
        # If a2l does not exist, create a new element for a2l
        if not a2l_is_found:
            screen.create_a2l_part(root_sub_element_item, list_temp, a2l_path_temp)
        tree.write(env_path_temp)

    @staticmethod
    def modify_a2l_part(item_sub_element, list_temp):
        """
        Modify existing a2l part
        :param item_sub_element: XML address of matched a2l
        :param list_temp: member list
        :return:
        """
        # Add all 'diff' element names to member list and remove them from XML
        for diff in item_sub_element.findall('diff'):
            list_temp.append(diff.get('name'))
            item_sub_element.remove(diff)
        # Remove duplicates
        list_temp = screen.remove_duplicates(list_temp)
        # Add all list elements to XML file with 'diff' tag
        for diff_new in list_temp:
            item_sub_element.append(ElT.Element(
                "diff",
                attrib={
                    "name": f"{diff_new}",
                    "READ_WRITE": "true"
                }))
            if diff_new == list_temp[-1]:
                item_sub_element[-1].tail = "\n" + " " * 8
            else:
                item_sub_element[-1].tail = "\n" + " " * 10

    @staticmethod
    def create_a2l_part(root_sub_element_item, list_temp, a2l_path_temp):
        """
        Create a2l part
        :param root_sub_element_item: Address of 'content' in XML
        :param list_temp: member list
        :param a2l_path_temp: a2l path
        :return:
        """
        # Get a2l name from a2l path
        a2l_name = os.path.basename(a2l_path_temp)
        for x in root_sub_element_item:
            if x.get('active'):
                # Deactivate last used a2l
                active_part = root_sub_element_item.find("./item[@active='true']")
                del active_part.attrib['active']
        # Create and activate new a2l
        # 'item' part
        new_element_item = ElT.Element('item', attrib={'active': 'true', 'name': f'{a2l_name}', 'type': 'asap2'})

        root_sub_element_item[-1].tail = '\n' + ' ' * 8
        new_element_item.text = '\n' + ' ' * 10
        new_element_item.tail = '\n' + ' ' * 6
        root_sub_element_item.append(new_element_item)
        # 'file' part
        new_element_file = ElT.Element('file')
        new_element_file.tail = '\n' + ' ' * 10
        new_element_file.text = f'{a2l_path_temp}'
        root_sub_element_item[-1].append(new_element_file)
        # 'commConfig' part
        new_element_comm_config = ElT.SubElement(new_element_item, 'commConfig',
                                                 attrib={'CanFd_BitRates': '',
                                                         'CanFd_Std': '1',
                                                         'Can_Btr': '28',
                                                         'Flx_HwId': '',
                                                         'HwId': 'J2534 [ES581]',
                                                         'Protocol': 'XcpOnCan'
                                                         })
        new_element_comm_config.tail = '\n' + ' ' * 10
        # 'diff' part
        for element in list_temp:
            new_element_diff = ElT.SubElement(new_element_item, 'diff',
                                              attrib={'READ_WRITE': 'true', 'name': f'{element}'})
            if list_temp[-1] != element:
                new_element_diff.tail = '\n' + ' ' * 10
            else:
                new_element_diff.tail = '\n' + ' ' * 8

    @staticmethod
    def message_box():
        """
        Pop up a message box that contains details of script
        :return:
        """
        # Get rid of unnecessary pop up window
        Tk().withdraw()
        return_value = messagebox.askokcancel("Parse Fts", "This program is about to parse FTS\n\n"
                                                           "Step 1 -> Select FTS\n"
                                                           "Step 2 -> Save XML File\n"
                                                           "Step 3 -> Select ENV file\n"
                                                           "Step 4 -> Select A2L file")
        # If cancelled
        if not return_value:
            print("User cancelled the pop up.")
            sys.exit()

    @staticmethod
    def info_box():
        """
        Pop up a info box
        :return:
        """
        # Get rid of unnecessary pop up window
        Tk().withdraw()
        messagebox.showinfo("FTS is parsed", "XML is created.\n"
                                             "ENV is modified.")


if __name__ == '__main__':
    screen = FTS()
    # Pop up a message box that give user the details
    screen.message_box()
    # Get FTS file path from user
    FTS_file_path = screen.select_file("Step 1 of 4     Select FTS document", ("Word documents", "*.docm"))
    # Parse FTS file and get the variables
    get_member_list = screen.reg(FTS_file_path)
    # Remove duplicate variables
    eliminated_member_list = screen.remove_duplicates(get_member_list)
    # Get the address to save xml file
    XML_file_path = screen.save_xml(FTS_file_path)
    # Create the XML file
    screen.write_file(eliminated_member_list, XML_file_path)
    # Get the ENV file path
    env_path = screen.select_file("Step 3 of 4     Select env document", ("Env Files", "*.env"))
    # Get the A2L file path
    a2l_path = screen.select_file("Step 4 of 4     Select A2L", ("A2L Files", "*.a2l"))
    # Check the A2L file, if it exist, modify it, if not, create it
    screen.check_a2l(env_path, a2l_path, eliminated_member_list)
    # Pop up info box that tells parsing is done
    screen.info_box()