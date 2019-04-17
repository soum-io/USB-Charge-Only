'''
By: Michael Shea
Email: mjshea3@illinois.edu
Phone: 708-203-8272
'''

# imports
from PyQt5 import QtCore, QtGui, QtWidgets, QtGui
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
import urllib.request
import os
import zipfile
import shutil
from shutil import copyfile
from subprocess import check_output, call
import sys
from functools import partial

# methods
'''
checks if current system is 64 bit or 32 bit - need this to see what
version of usbtreeview is used
'''
def is_windows_64bit():
    if 'PROCESSOR_ARCHITEW6432' in os.environ:
        return True
    return os.environ['PROCESSOR_ARCHITECTURE'].endswith('64')

'''
generate usb report
'''
def generateReport():
    # generate report of usb devices
    call("UsbTreeView -R=logs/usblog.txt", shell=True) # 0 if a-okay :D

    # need this to identify end of file line
    with open("logs/usblog.txt", "a", encoding="ascii") as myfile:
        myfile.write("\n==========")

'''
Parses generated USB report
'''
def parseReport():
    generateReport()
    # parse report
    log_loc = os.getcwd() + "\\logs\\usblog.txt"
    key_word = "== USB Port"

    # vars to keep track of what port we are currently parsing
    port_reports = dict()
    in_report = False
    cur_report_key = ""
    cur_report_data = ""
    # a computer can have multiple root hubs, each which can have multiple usb devices
    cur_root_hub = ""
    in_hub = False
    all_lines = list()

    with open(log_loc, "r") as log:
        # read through logs one at a time, and same them
        for line in log:
            # for some frustrating reason, there are a lot of null characters in the file.
            all_lines.append(line[:-1].translate({ord('\0'): None}))

    for line_idx, line in enumerate(all_lines):
        # var for loop control
        should_continue = False

        # checking if we are starting a new usb hub section
        if '= USB Root Hub =' in line and not in_report:
            in_hub = True
            continue

        # if we see a new section
        if '===' in line:
            # if we are currently looking at a port's information
            if in_report and (key_word in line or line_idx+1 == len(all_lines) or '= USB Root Hub =' in line):
                # check that if we see a port header - it is not a port that is part of an external usb hub
                # (we do not have support for external usb hubs at this time)
                stop = True
                if(line_idx+1!=len(all_lines) and '= USB Root Hub =' not in line):
                    port_desc_line = all_lines[line_idx+6]
                    # each port is described by a chain #a-#b-#c... where #a if root hub, #b is part of primary port,
                    # all other letters c, d, e... are external ports we do not look at currently
                    _, port_chain = port_desc_line.split(":")
                    port_chain = port_chain.strip()
                    # port chain goes past letter b if true
                    if(len(port_chain.split('-')) > 2):
                        stop = False

                if(stop):
                    # split up the data for the current port to be parsed more thoroughly
                    cur_report_data_lines = cur_report_data.split("\n")
                    # will hold all the ports data in a dict
                    cur_port_report = dict()
                    # will hold connected device data if port is in use
                    usb_device_report = dict()
                    # vars to keep track of what we have already seen
                    in_usb_device = False
                    usb_device_key = ""
                    usb_line = -1
                    seen_usb_device = False
                    seen_desc = ""
                    in_hub = False
                    # loop through lines that belong to current port
                    for idx, report_line in enumerate(cur_report_data_lines):
                        # if we get this section, it means that there is a device
                        # connected to the port, so start sorting through that data
                        if("USB Hub") in report_line and not seen_usb_device:
                            usb_line = idx
                            in_usb_device = True
                            seen_usb_device = True
                            seen_desc = "usb hub"
                        elif("USB Device") in report_line and not seen_usb_device:
                            usb_line = idx
                            in_usb_device = True
                            seen_usb_device = True
                            seen_desc = "usb device"
                        # if we are not sorting through conencted device info,
                        # then we are just sorting through general port information
                        elif ":" in report_line and not in_usb_device:
                            split_idx = report_line.find(":")
                            key = report_line[:split_idx].strip()
                            val = report_line[split_idx+1:].strip()
                            cur_port_report[key] = val
                        # if we started searching through connected device info, continue
                        # searching through it
                        elif in_usb_device and seen_desc == "usb device":
                            # dont want any useless blank lines
                            if len(report_line.strip()) == 0:
                                continue
                            else:
                                # if it starts with -- or ++, these are sections headers
                                # for connected device info. These will be keys in the usb
                                # device dict
                                if "--" in report_line or "++" in report_line:
                                    report_line = report_line.replace("-", "")
                                    report_line = report_line.replace("+", "")
                                    usb_device_key = report_line.strip()
                                    usb_device_report[usb_device_key] = dict()
                                # for each connected device header, parse each of its traits
                                # in key value pairs
                                else:
                                    split_idx = report_line.find(":")
                                    key = report_line[:split_idx].strip()
                                    val = report_line[split_idx+1:].strip()
                                    usb_device_report[usb_device_key][key] = val
                    # only save parsed information if "ConnectionIndex" is a key - I think this
                    # is an indicator that there is a phyical port in the computer and not just
                    # some port that could be connected to the cpu but is not. (otherwise we would
                    # have over 16 ports most of the time!)
                    if "ConnectionIndex" in cur_port_report:
                        # structure the port info depending on if it has a device connected to it or not
                        if(usb_line != -1 and seen_desc == "usb device"):
                            # make the text point to a brief text summary, but not about connected device
                            cur_port_report["text"] = "\n".join(cur_report_data_lines[:usb_line])
                            cur_port_report["connected_usb_info"] = usb_device_report
                        elif(usb_line != -1 and seen_desc == "usb hub"):
                            cur_port_report["text"] = cur_report_data
                            cur_port_report["connected_usb_info"] = "External USB Hub"
                        else:
                            cur_port_report["text"] = cur_report_data
                            cur_port_report["connected_usb_info"] = None
                        cur_report_key = cur_report_key.replace("=", "")
                        cur_report_key = cur_report_key.strip()
                        # just so it is the port number as the key
                        cur_report_key = cur_report_key[8:]
                        port_reports[cur_root_hub][cur_report_key] = cur_port_report
                    # reset vars
                    cur_report_data = ""
                    in_report = True
                    # if the new header we see is for a new port, init the key for a new dict entry
                    cur_report_key = line.strip()
                    should_continue = True

            # checking if we are starting a new usb hub section
            if('= USB Root Hub =' in line):
                in_hub = True
                in_report = False
                should_continue = True

            # will be true when we just saw start of new root hub and this is the first port's info of that hub
            if not in_report and key_word in line:
                port_desc_line = all_lines[line_idx+6]
                # each port is described by a chain #a-#b-#c... where #a if root hub, #b is part of primary port,
                # all other letters c, d, e... are external ports we do not look at currently
                _, port_chain = port_desc_line.split(":")
                port_chain = port_chain.strip()
                # port chain goes past letter b if true
                if not (len(port_chain.split('-')) > 2):
                    # reset vars
                    cur_report_data = ""
                    in_report = True
                    # if the new header we see is for a new port, init the key for a new dict entry
                    cur_report_key = line.strip()

        # if we did some operations above such that the rest of the loop would not make sense to run, continue
        if(should_continue):
            continue

        # see if we are in new hub section and came across hub name -> needed to enable/disable usb ports!
        if in_hub and line[:18] == "Device Description":
            in_hub = False
            _, hub_name = line.split(":")
            cur_root_hub = hub_name.strip()
            port_reports[cur_root_hub] = dict()
            continue

        # if we are looking at an ordinary line that is not a header, append it
        # to the current data for the current port we are looking at
        if in_report and line.strip() != '':
            if(cur_report_data == ''):
                cur_report_data = line
            else:
                cur_report_data += "\n" + line

    return port_reports

'''
find true number of physical ports and their port numbers
the idea of links here is that multiple ports can be the same physcial port (one is for usb3, the other is for usb2, etc...)
'''
def getPortNums():
    port_reports = parseReport()
    hub_usb_ports = dict()
    for hub in port_reports:
        # list of list, where each inner list is a list of port numbers that refer to a single physical port
        usb_ports = list()
        # keep track of ports we have already taken into account
        added_ports = set()
        for key in port_reports[hub]:
            companion_port = "-1"
            if "CompanionIndex" in port_reports[hub][key]:
                companion_port = port_reports[hub][key]["CompanionPortNumber"]
            # if there are multiple ports to the physical port
            if(companion_port != "-1"):
                # if we havent seen either port yet
                if companion_port not in added_ports and key not in added_ports:
                    usb_ports.append([key, companion_port])
                # if we havent seen the companion port yet
                elif companion_port not in added_ports:
                    for port in usb_ports:
                        if port[0] == key:
                            port.extend(companion_port)
                # if we havent seen the current port yet
                elif key not in added_ports:
                    for port in usb_ports:
                        if port[0] == companion_port:
                            port.extend(key)
                added_ports.add(key)
                added_ports.add(companion_port)
            # if there is just one port for the physical port
            else:
                # check if we have seen the port or not yet
                if key not in added_ports:
                    added_ports.add(key)
                    usb_ports.append([key])
        hub_usb_ports[hub] = usb_ports
    return hub_usb_ports, port_reports

'''
get usb support for each port
'''
def USBSupport():
    hub_usb_ports, port_reports = getPortNums()
    # main simplified usb port info for each physical port
    usb_dict = dict()
    for hub in hub_usb_ports:
        usb_dict[hub] = dict()
        for physical_port in hub_usb_ports[hub]:
            enabled = True
            usb_support = {3 : False, 2 : False, 1: False}
            # bool that states if a device is connected
            in_use = False
            device_name = "N/A"
            for port_num in physical_port:
                # check if port is in use -> this also determines how we
                # find supported usb protocols
                if port_reports[hub][port_num]["connected_usb_info"] != None and port_reports[hub][port_num]["connected_usb_info"] != "External USB Hub":
                    if("Device Information" in port_reports[hub][port_num]["connected_usb_info"] and "Problem Code" in port_reports[hub][port_num]["connected_usb_info"]["Device Information"]):
                        # 22 means the usb device is disabled
                        enabled = ("22" != port_reports[hub][port_num]["connected_usb_info"]["Device Information"]["Problem Code"][:2])
                    in_use = True
                    # port supports usb 1
                    if port_reports[hub][port_num]["connected_usb_info"]["Connection Information V2"]["Usb110"][0] == '1':
                        usb_support[1] = True
                    # port supports usb 2
                    if port_reports[hub][port_num]["connected_usb_info"]["Connection Information V2"]["Usb200"][0] == '1':
                        usb_support[2] = True
                    # port supports usb 3
                    if port_reports[hub][port_num]["connected_usb_info"]["Connection Information V2"]["Usb300"][0] == '1':
                        usb_support[3] = True
                    # get name of connected device
                    device_name = port_reports[hub][port_num]["connected_usb_info"]["Device Information"]['Device Description']
                elif port_reports[hub][port_num]["connected_usb_info"] != None and port_reports[hub][port_num]["connected_usb_info"] == "External USB Hub":
                    in_use = True
                    device_name = "External USB Hub"
                else:
                    # port supports usb 1
                    if "Usb110" in port_reports[hub][port_num] and port_reports[hub][port_num]["Usb110"][0] == '1':
                        usb_support[1] = True
                    # port supports usb 2
                    if "Usb200" in port_reports[hub][port_num] and port_reports[hub][port_num]["Usb200"][0] == '1':
                        usb_support[2] = True
                    # port supports usb 3
                    if "Usb300" in port_reports[hub][port_num] and port_reports[hub][port_num]["Usb300"][0] == '1':
                        usb_support[3] = True
            # each physical usb key will be the string composed of each of its port numbers
            # seperated by a comma
            usb_dict[hub][",".join(physical_port)] = dict()
            usb_dict[hub][",".join(physical_port)]["usb_support"] = usb_support
            usb_dict[hub][",".join(physical_port)]["in_use"] = in_use
            usb_dict[hub][",".join(physical_port)]["device_name"] = device_name
            usb_dict[hub][",".join(physical_port)]["enabled"] = enabled if in_use else False
    return usb_dict

'''
Start up script to get usb port info and download necessary files
'''
def startUp():
    # temp script to delete all files and start over
    if(os.path.isfile(os.getcwd() + "\\UsbTreeView.exe")):
        os.remove(os.getcwd() + "\\UsbTreeView.exe")
    if(os.path.isfile(os.getcwd() + "\\UsbTreeView.ini")):
        os.remove(os.getcwd() + "\\UsbTreeView.ini")
    if(os.path.isfile(os.getcwd() + "\\USBDeview.exe")):
        os.remove(os.getcwd() + "\\USBDeview.exe")
    if(os.path.isdir(os.getcwd() + "\\logs")):
        shutil.rmtree(os.getcwd() + "\\logs")

    # download usbtreeview and usbdeview
    if not os.path.isfile(os.getcwd() + "/UsbTreeView.exe"):
        urllib.request.urlretrieve('https://www.uwe-sieber.de/files/usbtreeview.zip', os.getcwd() + "/usbtreeview.zip")
        # unzip usbtreeview and make required directories
        with zipfile.ZipFile("usbtreeview.zip","r") as zip_ref:
            os.mkdir(os.getcwd()+"/usbtreeview")
            os.mkdir(os.getcwd()+"/logs")
            zip_ref.extractall(os.getcwd()+"/usbtreeview")
        # get if we are on windows 32 or 64, and move corresponding usbtreefile
        is_64 = is_windows_64bit()
        usbtreeview_loc = ""
        # url for program called USBDeview, allows us to disable usb devices so they can only do charge mode
        usbdeview_url = ""
        # get location of current usbtreeview application to use
        if(is_64):
            usbtreeview_loc = os.getcwd()+"\\usbtreeview\\x64\\UsbTreeView.exe"
            usbdeview_url = "https://www.nirsoft.net/utils/usbdeview-x64.zip"
        else:
            usbtreeview_loc = os.getcwd()+"\\usbtreeview\\win32\\UsbTreeView.exe"
            usbdeview_url = "https://www.nirsoft.net/utils/usbdeview.zip"

        # download and extract usbdeview
        urllib.request.urlretrieve(usbdeview_url, os.getcwd() + "/usbdeview.zip")
        with zipfile.ZipFile("usbdeview.zip","r") as zip_ref:
            os.mkdir(os.getcwd()+"/usbdview")
            zip_ref.extractall(os.getcwd()+"/usbdview")
        # copy the USBDeview.exe program to current location
        usbdeview_loc = os.getcwd()+"/usbdview/USBDeview.exe"
        copyfile(usbdeview_loc, os.getcwd() + "\\USBDeview.exe")

        # copy the tree view application to project directory
        copyfile(usbtreeview_loc, os.getcwd() + "\\UsbTreeView.exe")
        os.remove(os.getcwd() + "\\usbtreeview.zip")
        shutil.rmtree(os.getcwd() + "\\usbtreeview")
        os.remove(os.getcwd() + "\\usbdeview.zip")
        shutil.rmtree(os.getcwd() + "\\usbdview")


class Ui_USBControl(object):
    EXIT_CODE_REBOOT = -123

    def setupUi(self, USBControl, num_hubs, num_ports, hub_names, port_names, usb_dict):
        width_needed =  400*max(num_ports)
        height_needed =  23 + (500)/2*num_hubs
        USBControl.setObjectName("USBControl")
        USBControl.resize(width_needed, height_needed)
        font = QtGui.QFont()
        font.setPointSize(8)
        USBControl.setFont(font)
        self.centralwidget = QtWidgets.QWidget(USBControl)
        self.centralwidget.setObjectName("centralwidget")
        self.widget = QtWidgets.QWidget(self.centralwidget)
        self.widget.setGeometry(QtCore.QRect(0, 0, width_needed, height_needed))
        self.widget.setObjectName("widget")
        self.overall_VLayout = QtWidgets.QVBoxLayout(self.widget)
        self.overall_VLayout.setContentsMargins(0, 0, 0, 0)
        self.overall_VLayout.setObjectName("overall_VLayout")
        self.refresh_btn = QtWidgets.QPushButton(self.widget)
        self.refresh_btn.setObjectName("refresh_btn")
        self.refresh_btn.clicked.connect(partial(self.refresh_btn_clicked, num_hubs, num_ports, hub_names, port_names))
        self.overall_VLayout.addWidget(self.refresh_btn)
        self.hubs_HLayout = QtWidgets.QVBoxLayout()
        self.hubs_HLayout.setObjectName("hubs_HLayout")


        self.portFrames = list()
        self.USBHUB3_VLayout = list()
        self.hubName1_label = list()
        self.USBHUB3_HLayout = list()
        self.port1_vlayout = list()
        self.port1_topVLayout = list()
        self.hubName1_2 =  list()
        self.port1Change = list()
        self.port1Details = list()
        self.port1_infoVLayout = list()
        self.pot1_usb3VLayout = list()
        self.port1usb3_label = list()
        self.port1usb3_answer = list()
        self.pot1_usb2VLayout = list()
        self.port1usb2_label = list()
        self.port1usb2_answer = list()
        self.pot1_usb1VLayout = list()
        self.port1usb1_label = list()
        self.port1usb1_answer = list()
        self.pot1_inUseVLayout = list()
        self.port1InUse_label = list()
        self.port1InUse_answer = list()
        self.pot1_deviceVLayout = list()
        self.port1DeviceName_label = list()
        self.port1DeviceName_answer = list()
        self.pot1_chargeVLayout = list()
        self.port1ChargeOnly_label = list()
        self.port1ChargeOnly_answer = list()
        self.port2_vlayout_2 = list()
        self.port2_topVLayout_2 = list()
        self.hubName1_4 = list()
        self.port1Change_3 = list()
        self.port1Details_3 = list()
        self.port2_infoVLayout_2 = list()
        self.pot2_usb3VLayout_2 = list()
        self.port1usb3_label_3 = list()
        self.port1usb3_answer_3 = list()
        self.pot2_usb2VLayout_2 = list()
        self.port1usb2_label_3 = list()
        self.port1usb2_answer_3 = list()
        self.pot2_usb1VLayout_2 = list()
        self.port1usb1_label_3 = list()
        self.port1usb1_answer_3 = list()
        self.pot2_inUseVLayout_2 = list()
        self.port1InUse_label_3 = list()
        self.port1InUse_answer_3 = list()
        self.pot2_deviceVLayout_2 = list()
        self.port1DeviceName_label_3 = list()
        self.port1DeviceName_answer_3 = list()
        self.pot2_chargeVLayout_2 = list()
        self.port1ChargeOnly_label_3 = list()
        self.port1ChargeOnly_answer_3 = list()


        for hub in range(num_hubs):
            self.USBHUB3_VLayout.append(QtWidgets.QVBoxLayout())
            self.USBHUB3_VLayout[hub].setObjectName("USBHUB3_VLayout")
            self.hubName1_label.append(QtWidgets.QLabel(self.widget))
            font = QtGui.QFont()
            font.setPointSize(13)
            font.setBold(True)
            font.setWeight(75)
            self.hubName1_label[hub].setFont(font)
            self.hubName1_label[hub].setFrameShape(QtWidgets.QFrame.NoFrame)
            self.hubName1_label[hub].setLineWidth(1)
            self.hubName1_label[hub].setObjectName("hubName1_label")
            self.hubName1_label[hub].setAlignment(Qt.AlignCenter)
            self.USBHUB3_VLayout[hub].addWidget(self.hubName1_label[hub])
            self.USBHUB3_HLayout.append(QtWidgets.QHBoxLayout())
            self.USBHUB3_HLayout[hub].setObjectName("USBHUB3_HLayout")


            self.portFrames.append(list())
            self.port1_vlayout.append(list())
            self.port1_topVLayout.append(list())
            self.hubName1_2.append(list())
            self.port1Change.append(list())
            self.port1Details.append(list())
            self.port1_infoVLayout.append(list())
            self.pot1_usb3VLayout.append(list())
            self.port1usb3_label.append(list())
            self.port1usb3_answer.append(list())
            self.pot1_usb2VLayout.append(list())
            self.port1usb2_label.append(list())
            self.port1usb2_answer.append(list())
            self.pot1_usb1VLayout.append(list())
            self.port1usb1_label.append(list())
            self.port1usb1_answer.append(list())
            self.pot1_inUseVLayout.append(list())
            self.port1InUse_label.append(list())
            self.port1InUse_answer.append(list())
            self.pot1_deviceVLayout.append(list())
            self.port1DeviceName_label.append(list())
            self.port1DeviceName_answer.append(list())
            self.pot1_chargeVLayout.append(list())
            self.port1ChargeOnly_label.append(list())
            self.port1ChargeOnly_answer.append(list())


            for port in range(num_ports[hub]):
                self.portFrames[hub].append(QtWidgets.QFrame()) # create the frame object so it can be collapsed
                self.port1_vlayout[hub].append(QtWidgets.QVBoxLayout())
                self.port1_vlayout[hub][port].setObjectName("port1_vlayout")
                self.port1_vlayout[hub][port].setAlignment(QtCore.Qt.AlignTop)
                self.port1_topVLayout[hub].append(QtWidgets.QHBoxLayout())
                self.port1_topVLayout[hub][port].setObjectName("port1_topVLayout")
                self.hubName1_2[hub].append(QtWidgets.QLabel(self.widget))
                font = QtGui.QFont()
                font.setPointSize(10)
                font.setBold(True)
                font.setUnderline(True)
                font.setWeight(75)
                self.hubName1_2[hub][port].setFont(font)
                self.hubName1_2[hub][port].setFrameShape(QtWidgets.QFrame.NoFrame)
                self.hubName1_2[hub][port].setLineWidth(1)
                self.hubName1_2[hub][port].setObjectName("hubName1_2")
                self.port1_topVLayout[hub][port].addWidget(self.hubName1_2[hub][port])
                self.port1Change[hub].append(QtWidgets.QPushButton(self.widget))
                self.port1Change[hub][port].setObjectName("port1Change")
                self.port1Change[hub][port].clicked.connect(partial(self.changePortMode, str(usb_dict[hub_names[hub]][port_names[hub][port]]["device_name"]), hub, port, usb_dict, num_hubs, num_ports, hub_names, port_names))
                self.port1_topVLayout[hub][port].addWidget(self.port1Change[hub][port])
                self.port1Details[hub].append(QtWidgets.QPushButton(self.widget))
                self.port1Details[hub][port].setObjectName("port1Details")
                self.port1Details[hub][port].clicked.connect(partial(self.detailView, hub, port))
                self.port1_topVLayout[hub][port].addWidget(self.port1Details[hub][port])
                self.port1_vlayout[hub][port].addLayout(self.port1_topVLayout[hub][port])
                self.port1_infoVLayout[hub].append(QtWidgets.QVBoxLayout())
                self.port1_infoVLayout[hub][port].setObjectName("port1_infoVLayout")
                self.pot1_usb3VLayout[hub].append(QtWidgets.QHBoxLayout())
                self.pot1_usb3VLayout[hub][port].setObjectName("pot1_usb3VLayout")
                self.port1usb3_label[hub].append(QtWidgets.QLabel(self.widget))
                font = QtGui.QFont()
                font.setPointSize(10)
                font.setBold(False)
                font.setUnderline(False)
                font.setWeight(50)
                self.port1usb3_label[hub][port].setFont(font)
                self.port1usb3_label[hub][port].setFrameShape(QtWidgets.QFrame.NoFrame)
                self.port1usb3_label[hub][port].setLineWidth(1)
                self.port1usb3_label[hub][port].setObjectName("port1usb3_label")
                self.pot1_usb3VLayout[hub][port].addWidget(self.port1usb3_label[hub][port])
                self.port1usb3_answer[hub].append(QtWidgets.QLabel(self.widget))
                font = QtGui.QFont()
                font.setPointSize(10)
                font.setBold(False)
                font.setUnderline(False)
                font.setWeight(50)
                self.port1usb3_answer[hub][port].setFont(font)
                self.port1usb3_answer[hub][port].setFrameShape(QtWidgets.QFrame.NoFrame)
                self.port1usb3_answer[hub][port].setLineWidth(1)
                self.port1usb3_answer[hub][port].setObjectName("port1usb3_answer")
                self.pot1_usb3VLayout[hub][port].addWidget(self.port1usb3_answer[hub][port])
                self.port1_infoVLayout[hub][port].addLayout(self.pot1_usb3VLayout[hub][port])
                self.pot1_usb2VLayout[hub].append(QtWidgets.QHBoxLayout())
                self.pot1_usb2VLayout[hub][port].setObjectName("pot1_usb2VLayout")
                self.port1usb2_label[hub].append(QtWidgets.QLabel(self.widget))
                font = QtGui.QFont()
                font.setPointSize(10)
                font.setBold(False)
                font.setUnderline(False)
                font.setWeight(50)
                self.port1usb2_label[hub][port].setFont(font)
                self.port1usb2_label[hub][port].setFrameShape(QtWidgets.QFrame.NoFrame)
                self.port1usb2_label[hub][port].setLineWidth(1)
                self.port1usb2_label[hub][port].setObjectName("port1usb2_label")
                self.pot1_usb2VLayout[hub][port].addWidget(self.port1usb2_label[hub][port])
                self.port1usb2_answer[hub].append(QtWidgets.QLabel(self.widget))
                font = QtGui.QFont()
                font.setPointSize(10)
                font.setBold(False)
                font.setUnderline(False)
                font.setWeight(50)
                self.port1usb2_answer[hub][port].setFont(font)
                self.port1usb2_answer[hub][port].setFrameShape(QtWidgets.QFrame.NoFrame)
                self.port1usb2_answer[hub][port].setLineWidth(1)
                self.port1usb2_answer[hub][port].setObjectName("port1usb2_answer")
                self.pot1_usb2VLayout[hub][port].addWidget(self.port1usb2_answer[hub][port])
                self.port1_infoVLayout[hub][port].addLayout(self.pot1_usb2VLayout[hub][port])
                self.pot1_usb1VLayout[hub].append(QtWidgets.QHBoxLayout())
                self.pot1_usb1VLayout[hub][port].setObjectName("pot1_usb1VLayout")
                self.port1usb1_label[hub].append(QtWidgets.QLabel(self.widget))
                font = QtGui.QFont()
                font.setPointSize(10)
                font.setBold(False)
                font.setUnderline(False)
                font.setWeight(50)
                self.port1usb1_label[hub][port].setFont(font)
                self.port1usb1_label[hub][port].setFrameShape(QtWidgets.QFrame.NoFrame)
                self.port1usb1_label[hub][port].setLineWidth(1)
                self.port1usb1_label[hub][port].setObjectName("port1usb1_label")
                self.pot1_usb1VLayout[hub][port].addWidget(self.port1usb1_label[hub][port])
                self.port1usb1_answer[hub].append(QtWidgets.QLabel(self.widget))
                font = QtGui.QFont()
                font.setPointSize(10)
                font.setBold(False)
                font.setUnderline(False)
                font.setWeight(50)
                self.port1usb1_answer[hub][port].setFont(font)
                self.port1usb1_answer[hub][port].setFrameShape(QtWidgets.QFrame.NoFrame)
                self.port1usb1_answer[hub][port].setLineWidth(1)
                self.port1usb1_answer[hub][port].setObjectName("port1usb1_answer")
                self.pot1_usb1VLayout[hub][port].addWidget(self.port1usb1_answer[hub][port])
                self.port1_infoVLayout[hub][port].addLayout(self.pot1_usb1VLayout[hub][port])
                self.pot1_inUseVLayout[hub].append(QtWidgets.QHBoxLayout())
                self.pot1_inUseVLayout[hub][port].setObjectName("pot1_inUseVLayout")
                self.port1InUse_label[hub].append(QtWidgets.QLabel(self.widget))
                font = QtGui.QFont()
                font.setPointSize(10)
                font.setBold(False)
                font.setUnderline(False)
                font.setWeight(50)
                self.port1InUse_label[hub][port].setFont(font)
                self.port1InUse_label[hub][port].setFrameShape(QtWidgets.QFrame.NoFrame)
                self.port1InUse_label[hub][port].setLineWidth(1)
                self.port1InUse_label[hub][port].setObjectName("port1InUse_label")
                self.pot1_inUseVLayout[hub][port].addWidget(self.port1InUse_label[hub][port])
                self.port1InUse_answer[hub].append(QtWidgets.QLabel(self.widget))
                font = QtGui.QFont()
                font.setPointSize(10)
                font.setBold(False)
                font.setUnderline(False)
                font.setWeight(50)
                self.port1InUse_answer[hub][port].setFont(font)
                self.port1InUse_answer[hub][port].setFrameShape(QtWidgets.QFrame.NoFrame)
                self.port1InUse_answer[hub][port].setLineWidth(1)
                self.port1InUse_answer[hub][port].setObjectName("port1InUse_answer")
                self.pot1_inUseVLayout[hub][port].addWidget(self.port1InUse_answer[hub][port])
                self.port1_infoVLayout[hub][port].addLayout(self.pot1_inUseVLayout[hub][port])
                self.pot1_deviceVLayout[hub].append(QtWidgets.QHBoxLayout())
                self.pot1_deviceVLayout[hub][port].setObjectName("pot1_deviceVLayout")
                self.port1DeviceName_label[hub].append(QtWidgets.QLabel(self.widget))
                font = QtGui.QFont()
                font.setPointSize(10)
                font.setBold(False)
                font.setUnderline(False)
                font.setWeight(50)
                self.port1DeviceName_label[hub][port].setFont(font)
                self.port1DeviceName_label[hub][port].setFrameShape(QtWidgets.QFrame.NoFrame)
                self.port1DeviceName_label[hub][port].setLineWidth(1)
                self.port1DeviceName_label[hub][port].setObjectName("port1DeviceName_label")
                self.pot1_deviceVLayout[hub][port].addWidget(self.port1DeviceName_label[hub][port])
                self.port1DeviceName_answer[hub].append(QtWidgets.QLabel(self.widget))
                font = QtGui.QFont()
                font.setPointSize(10)
                font.setBold(False)
                font.setUnderline(False)
                font.setWeight(50)
                self.port1DeviceName_answer[hub][port].setFont(font)
                self.port1DeviceName_answer[hub][port].setFrameShape(QtWidgets.QFrame.NoFrame)
                self.port1DeviceName_answer[hub][port].setLineWidth(1)
                self.port1DeviceName_answer[hub][port].setObjectName("port1DeviceName_answer")
                self.pot1_deviceVLayout[hub][port].addWidget(self.port1DeviceName_answer[hub][port])
                self.port1_infoVLayout[hub][port].addLayout(self.pot1_deviceVLayout[hub][port])
                self.pot1_chargeVLayout[hub].append(QtWidgets.QHBoxLayout())
                self.pot1_chargeVLayout[hub][port].setObjectName("pot1_chargeVLayout")
                self.port1ChargeOnly_label[hub].append(QtWidgets.QLabel(self.widget))
                font = QtGui.QFont()
                font.setPointSize(10)
                font.setBold(False)
                font.setUnderline(False)
                font.setWeight(50)
                self.port1ChargeOnly_label[hub][port].setFont(font)
                self.port1ChargeOnly_label[hub][port].setFrameShape(QtWidgets.QFrame.NoFrame)
                self.port1ChargeOnly_label[hub][port].setLineWidth(1)
                self.port1ChargeOnly_label[hub][port].setObjectName("port1ChargeOnly_label")
                self.pot1_chargeVLayout[hub][port].addWidget(self.port1ChargeOnly_label[hub][port])
                self.port1ChargeOnly_answer[hub].append(QtWidgets.QLabel(self.widget))
                font = QtGui.QFont()
                font.setPointSize(10)
                font.setBold(False)
                font.setUnderline(False)
                font.setWeight(50)
                self.port1ChargeOnly_answer[hub][port].setFont(font)
                self.port1ChargeOnly_answer[hub][port].setFrameShape(QtWidgets.QFrame.NoFrame)
                self.port1ChargeOnly_answer[hub][port].setLineWidth(1)
                self.port1ChargeOnly_answer[hub][port].setObjectName("port1ChargeOnly_answer")
                self.pot1_chargeVLayout[hub][port].addWidget(self.port1ChargeOnly_answer[hub][port])
                self.port1_infoVLayout[hub][port].addLayout(self.pot1_chargeVLayout[hub][port])
                self.portFrames[hub][port].setLayout(self.port1_infoVLayout[hub][port])
                # self.port1_vlayout[hub][port].addLayout(self.port1_infoVLayout[hub][port])
                self.port1_vlayout[hub][port].addWidget(self.portFrames[hub][port])
                margin = 25
                if(port == 0):
                    self.port1_vlayout[hub][port].setContentsMargins(0, 0, margin, 0)
                elif(port == num_ports[hub]-1):
                    self.port1_vlayout[hub][port].setContentsMargins(margin, 0, 0, 0)
                else:
                    self.port1_vlayout[hub][port].setContentsMargins(margin, 0, margin, 0)

                self.USBHUB3_HLayout[hub].addLayout(self.port1_vlayout[hub][port])


            self.USBHUB3_VLayout[hub].addLayout(self.USBHUB3_HLayout[hub])
            margin = 25
            if(hub == 0):
                self.USBHUB3_VLayout[hub].setContentsMargins(0, 0, 0, margin)
            elif(hub == num_hubs-1):
                self.USBHUB3_VLayout[hub].setContentsMargins(0, margin, 0, 0)
            else:
                self.USBHUB3_VLayout[hub].setContentsMargins(0, margin, 0, margin)

            self.hubs_HLayout.addLayout(self.USBHUB3_VLayout[hub])


        self.overall_VLayout.addLayout(self.hubs_HLayout)
        USBControl.setCentralWidget(self.centralwidget)
        self.statusbar = QtWidgets.QStatusBar(USBControl)
        self.statusbar.setObjectName("statusbar")
        USBControl.setStatusBar(self.statusbar)

        self.retranslateUi(USBControl, num_hubs, num_ports, hub_names, port_names, usb_dict)
        QtCore.QMetaObject.connectSlotsByName(USBControl)

    def retranslateUi(self, USBControl, num_hubs, num_ports, hub_names, port_names, usb_dict):
        _translate = QtCore.QCoreApplication.translate
        USBControl.setWindowTitle(_translate("USBControl", "MainWindow"))
        self.translateUi_helper(num_hubs, num_ports, hub_names, port_names, usb_dict)

    def translateUi_helper(self, num_hubs, num_ports, hub_names, port_names, usb_dict):
        _translate = QtCore.QCoreApplication.translate
        self.refresh_btn.setText(_translate("USBControl", "Refresh"))
        for hub in range(num_hubs):
            self.hubName1_label[hub].setText(_translate("USBControl", hub_names[hub]))
            for port in range(num_ports[hub]):
                self.hubName1_2[hub][port].setText(_translate("USBControl", "USB Port #" + str(port+1)))
                self.port1Details[hub][port].setText(_translate("USBControl", "Hide Details"))
                self.port1usb3_label[hub][port].setText(_translate("USBControl", "USB 3 Support:"))
                self.port1usb3_answer[hub][port].setText(_translate("USBControl", str(usb_dict[hub_names[hub]][port_names[hub][port]]["usb_support"][3])))
                self.port1usb2_label[hub][port].setText(_translate("USBControl", "USB 2 Support:"))
                self.port1usb2_answer[hub][port].setText(_translate("USBControl", str(usb_dict[hub_names[hub]][port_names[hub][port]]["usb_support"][2])))
                self.port1usb1_label[hub][port].setText(_translate("USBControl", "USB 1 Support:"))
                self.port1usb1_answer[hub][port].setText(_translate("USBControl", str(usb_dict[hub_names[hub]][port_names[hub][port]]["usb_support"][1])))
                self.port1InUse_label[hub][port].setText(_translate("USBControl", "In Use:"))
                self.port1InUse_answer[hub][port].setText(_translate("USBControl", str(usb_dict[hub_names[hub]][port_names[hub][port]]["in_use"])))
                self.port1DeviceName_label[hub][port].setText(_translate("USBControl", "Device Name:"))
                self.port1DeviceName_answer[hub][port].setText(_translate("USBControl", str(usb_dict[hub_names[hub]][port_names[hub][port]]["device_name"])[:20]))
                self.port1ChargeOnly_label[hub][port].setText(_translate("USBControl", "Charge Only:"))
                if(str(usb_dict[hub_names[hub]][port_names[hub][port]]["device_name"]) != "External USB Hub"):
                    if(usb_dict[hub_names[hub]][port_names[hub][port]]["in_use"] == False):
                        self.port1ChargeOnly_answer[hub][port].setText(_translate("USBControl", "N/A"))
                        # cant enable or disable a device if non are plugged in!
                        self.port1Change[hub][port].setEnabled(False)
                        self.port1Change[hub][port].setText(_translate("USBControl", "Turn Charge Only On"))
                    else:
                        if(not usb_dict[hub_names[hub]][port_names[hub][port]]["enabled"]):
                            self.port1Change[hub][port].setText(_translate("USBControl", "Turn Charge Only Off"))
                        else:
                            self.port1Change[hub][port].setText(_translate("USBControl", "Turn Charge Only On"))
                        self.port1Change[hub][port].setEnabled(True)
                        self.port1ChargeOnly_answer[hub][port].setText(_translate("USBControl", str(not usb_dict[hub_names[hub]][port_names[hub][port]]["enabled"])))
                else: # external hubs not currently supported
                    self.port1ChargeOnly_answer[hub][port].setText(_translate("USBControl", "Not Supported for Hubs"))
                    self.port1Change[hub][port].setText(_translate("USBControl", "Charge Only Not Allowed"))
                    self.port1Change[hub][port].setEnabled(False)
                    self.port1usb3_answer[hub][port].setText(_translate("USBControl", "N/A"))
                    self.port1usb2_answer[hub][port].setText(_translate("USBControl", "N/A"))
                    self.port1usb1_answer[hub][port].setText(_translate("USBControl", "N/A"))



    '''
    refresh gui with to check for new inputs
    '''
    def refresh_btn_clicked(self, num_hubs, num_ports, hub_names, port_names):
        usb_dict = USBSupport()
        # QtWidgets.QApplication.exit(self.EXIT_CODE_REBOOT)
        self.translateUi_helper(num_hubs, num_ports, hub_names, port_names, usb_dict)

    '''
    show or hide a ports details
    '''
    def detailView(self, hub, port):
        _translate = QtCore.QCoreApplication.translate
        if(not self.portFrames[hub][port].isHidden()):
            self.portFrames[hub][port].hide()
            self.port1Details[hub][port].setText(_translate("USBControl", "Show Details"))
        else:
            self.port1Details[hub][port].setText(_translate("USBControl", "Hide Details"))
            self.portFrames[hub][port].show()
    '''
    descide which mode to change to
    '''
    def changePortMode(self, desc, hub, port, usb_dict, num_hubs, num_ports, hub_names, port_names):
        usb_dict = USBSupport()
        chargeOnly = (not usb_dict[hub_names[hub]][port_names[hub][port]]["enabled"])
        if(chargeOnly):
            self.chargeOnlyOff(desc)
        else:
            self.chargeOnlyOn(desc)
        # make sure we wait until the correct change is applied
        # try 3 times max
        count = 0
        while(True):
            count += 1
            usb_dict = USBSupport()
            if((not usb_dict[hub_names[hub]][port_names[hub][port]]["enabled"] != chargeOnly) or count == 3):
                if(count == 3):
                    err = QMessageBox()
                    err.setText("Operation Failed. Please try again.")
                    err.exec_()

                break

        # refresh to show updates
        self.refresh_btn_clicked(num_hubs, num_ports, hub_names, port_names)

    '''
    enable data transfer for usb device with the provided description
    '''
    def chargeOnlyOff(self, device_desc):
        call('USBDeview.exe /RunAsAdmin /enable "' + device_desc+'"' , shell=True) # 0 if a-okay :D

    '''
    disable data transfer for usb device with the provided description
    '''
    def chargeOnlyOn(self, device_desc):
        call('USBDeview.exe /RunAsAdmin /disable "' + device_desc+'"' , shell=True) # 0 if a-okay :D






if __name__ == "__main__":
    currentExitCode = Ui_USBControl.EXIT_CODE_REBOOT
    startUp()
    usb_dict = USBSupport()
    while currentExitCode == Ui_USBControl.EXIT_CODE_REBOOT:
        # number of hubs in computer
        num_hubs = len(list(usb_dict))
        # name of each hub in computer
        hub_names = list(usb_dict)
        # how many ports each hub has
        num_ports = [];
        # port names of each hub
        port_names = [];
        for hub_name in hub_names:
            num_ports.append(len(list(usb_dict[hub_name])))
            hub_port_names = []
            for port in usb_dict[hub_name]:
                hub_port_names.append(port)
            port_names.append(hub_port_names)
        # start gui logic
        app = QtWidgets.QApplication(sys.argv)
        USBControl = QtWidgets.QMainWindow()
        ui = Ui_USBControl()
        ui.setupUi(USBControl, num_hubs, num_ports, hub_names, port_names, usb_dict)
        USBControl.show()
        currentExitCode = app.exec_()
        app = None
        # sys.exit(app.exec_())
