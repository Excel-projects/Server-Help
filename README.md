[<img align="left" src="Images/ReadMe/App.png" width="64px" >](https://github.com/aduguid/ServerActions/blob/master/VBA/ServerActions.xlsm?raw=true "Download File")
# Server Actions  <span class="Application_Version">2.0.0.0</span>
This is an Excel Addin written in VBA. It allows the user to use an Excel table to ping a list of servers and create a file for Microsoft Remote Desktop Manager. This is used for quickly determining which servers are offline in a list.
<h1 align="center">
  <img src="Images/ReadMe/ServerActionsExample2.gif" alt="vbaping" />
</h1>

## Table of Contents
- <a href="#dependencies">Dependencies</a>
- <a href="#glossary-of-terms">Glossary of Terms</a>
- <a href="#functionality">Functionality</a>
    - <a href="#format-data-table">Format Data Table</a>
        - <a href="#format-as-table">Format as Table</a>
        - <a href="#freeze-panes">Freeze Panes</a>
        - <a href="#remove-duplicates">Remove Duplicates</a>
    - <a href="#ping-test">Ping Test</a>
        - <a href="#ping">Ping</a>
        - <a href="#server-column">Server Column</a>
        - <a href="#ping-column">Ping Column</a>
    - <a href="#remote-desktop-manager">Remote Desktop Manager</a>
        - <a href="#create-file">Create File </a>
        - <a href="#server">Server</a>
        - <a href="#description">Description</a>
        - <a href="#file-name">File Name</a>
    - <a href="#options">Options</a>
        - <a href="#refresh-lists">Refresh Lists</a>
        - <a href="#visual-basic">Visual Basic</a>
    - <a href="#about">About</a>

<a id="user-content-dependencies" class="anchor" href="#dependencies" aria-hidden="true"> </a>
## Dependencies
|Software                        |Dependency                 |
|:-------------------------------|:--------------------------|
|[Visual Basic for Applications](https://msdn.microsoft.com/en-us/vba/vba-language-reference)|Code|
|[Extensible Markup Language (XML)](https://www.rondebruin.nl/win/s2/win001.htm)|Ribbon|
|[Remote Desktop Manager](https://www.microsoft.com/en-au/download/details.aspx?id=44989)|Export File|
|[ScreenToGif](http://www.screentogif.com/)|Read Me|
|[Snagit](http://discover.techsmith.com/snagit-non-brand-desktop/?gclid=CNzQiOTO09UCFVoFKgod9EIB3g)|Read Me|

<a id="user-content-glossary-of-terms" class="anchor" href="#glossary-of-terms" aria-hidden="true"> </a>
## Glossary of Terms

| Term                      | Meaning                                                                                  |
|:--------------------------|:-----------------------------------------------------------------------------------------|
| Ping |Ping is a computer network administration software utility used to test the reachability of a host on an Internet Protocol (IP) network. It measures the round-trip time for messages sent from the originating host to a destination computer that are echoed back to the source. Ping operates by sending Internet Control Message Protocol (ICMP/ICMP6 ) Echo Request packets to the target host and waiting for an ICMP Echo Reply. The program reports errors, packet loss, and a statistical summary of the results, typically including the minimum, maximum, the mean round-trip times, and standard deviation of the mean. The command-line options of the ping utility and its output vary between the numerous implementations. Options may include the size of the payload, count of tests, limits for the number of network hops (TTL) that probes traverse, and interval between the requests. Many systems provide a companion utility ping6, for testing on Internet Protocol version 6 (IPv6) networks. |
| VBA |Visual Basic for Applications (VBA) is an implementation of Microsoft's event-driven programming language Visual Basic 6 and uses the Visual Basic Runtime Library. However, VBA code normally can only run within a host application, rather than as a standalone program. VBA can, however, control one application from another using OLE Automation. VBA can use, but not create, ActiveX/COM DLLs, and later versions add support for class modules.|
| XML |Extensible Markup Language (XML) is a markup language that defines a set of rules for encoding documents in a format that is both human-readable and machine-readable. The design goals of XML emphasize simplicity, generality, and usability across the Internet. It is a textual data format with strong support via Unicode for different human languages. Although the design of XML focuses on documents, the language is widely used for the representation of arbitrary data structures such as those used in web services.|

<a id="user-content-functionality" class="anchor" href="#functionality" aria-hidden="true"> </a>
## Functionality
This Excel ribbon named “Server Actions” is inserted after the “Home” tab when Excel opens.  Listed below is the detailed functionality of this application and its components.  

<a id="user-content-format-data-table" class="anchor" href="#format-data-table" aria-hidden="true"> </a>
### Format Data Table (Group)
These buttons have the following constraints: 
* Only runs on visible columns/rows. 

<a id="user-content-format-as-table" class="anchor" href="#format-as-table" aria-hidden="true"> </a>
####	Format as Table (Button)
* Quickly format a range of cells and convert it to a Table by choosing a pre-defined Table Style. 

<a id="user-content-freeze-panes" class="anchor" href="#freeze-panes" aria-hidden="true"> </a>
####	Freeze Panes (Button)
* Keep a portion of the sheet visible while the rest of the sheet scrolls
* Defaults to invisible from the install

<a id="user-content-remove-duplicates" class="anchor" href="#remove-duplicates" aria-hidden="true"> </a>
#### Remove Duplicates (Button)
* Delete duplicate rows from a sheet
* Defaults to invisible from the install

<a id="user-content-ping-test" class="anchor" href="#ping-test" aria-hidden="true"> </a>
###	Ping Test (Group)

<a id="user-content-ping" class="anchor" href="#ping" aria-hidden="true"> </a>
####	Ping (Button)
* This will ping the visible servers in the active table.

<a id="user-content-server-column" class="anchor" href="#server-column" aria-hidden="true"> </a>
####	Server Column (Dropdown)
* A list of column names from the active table.

<a id="user-content-ping-column" class="anchor" href="#ping-column" aria-hidden="true"> </a>
####	Ping Column (Dropdown)
* A list of column names from the active table. If the column doesn't exist, it will be created.

<a id="user-content-remote-desktop-manager" class="anchor" href="#remote-desktop-manager" aria-hidden="true"> </a>
###	Remote Desktop Manager (Group)

<a id="user-content-create-file" class="anchor" href="#create-file" aria-hidden="true"> </a>
####	Create File (Button)
* Creates a Remote Desktop Manager file of the active table list of servers

<a id="user-content-server" class="anchor" href="#server" aria-hidden="true"> </a>
####	Server (Dropdown)
* A list of column names from the active table.

<a id="user-content-description" class="anchor" href="#description" aria-hidden="true"> </a>
####	Description (Dropdown)
* A list of column names from the active table.

<a id="user-content-file-name" class="anchor" href="#file-name" aria-hidden="true"> </a>
####	File Name (Textbox)
* The file name to save the list of servers for Remote Desktop Manager.

<a id="user-content-options" class="anchor" href="#options" aria-hidden="true"> </a>
###	Options (Group)

<a id="user-content-refresh-lists" class="anchor" href="#refresh-lists" aria-hidden="true"> </a>
####	Refresh Lists (Button)
* Refreshes all the dropdown values from the active table column names.

<a id="user-content-visual-basic" class="anchor" href="#visual-basic" aria-hidden="true"> </a>
####	Visual Basic (Button)
* Opens the Visual Basic editor.

<a id="user-content-about" class="anchor" href="#about" aria-hidden="true"> </a>
###	About (Group)

#### Description (Label)
* The application name with the version

#### Install Date (Label)
* The install date of the application

#### Copyright (Label)
* The author’s name
