<html>
<head>
</head>

<img align="left" src="Images/ReadMe/App.png" width="64px" >

# Server Actions  <span class="Application_Version">2.0.0.0</span> [<img src="Images/ReadMe/VBA.png" width="80">](https://github.com/aduguid/ServerActions/blob/master/VBA/ServerActions.xlsm?raw=true  "Download the VBA Add-In")
This is an Excel Add-In written in Visual Studio Community 2017 C#/VB.NET and VBA. It allows the user to use an Excel table to ping a list of servers and create a file for Microsoft Remote Desktop Manager. This is used for quickly determining which servers are offline in a list. This has now been tested on Windows (7, 8.1, 10) & Excel (2010, 2013, 2016).
<h1 align="center">
  <img src="Images/ReadMe/ServerActionsExample2.0.gif" alt="vbaping" />
</h1>

## Table of Contents
- <a href="#dependencies">Dependencies</a>
- <a href="#glossary-of-terms">Glossary of Terms</a>
- <a href="#functionality">Functionality</a>
    - <a href="#ping-test">Ping Test</a>
        - <a href="#ping">Ping (button)</a>
        - <a href="#server-column">Server (dropdown)</a>
        - <a href="#ping-column">Ping (dropdown)</a>
    - <a href="#remote-desktop-manager">Remote Desktop Manager</a>
        - <a href="#create-file">Create File (button)</a>
        - <a href="#server">Server (dropdown)</a>
        - <a href="#description">Description (dropdown)</a>
        - <a href="#comment">Comment (dropdown)</a>
        - <a href="#group">Group (dropdown)</a>
        - <a href="#file-name">File Name (textbox)</a>
    - <a href="#options">Options</a>
        - <a href="#rebuild-list">Rebuild Server List (button)</a>
        - <a href="#refresh-lists">Refresh Dropdowns (button)</a>
        - <a href="#settings">Add-In Settings (button)</a>
        - <a href="#visual-basic">Visual Basic (button)</a>
    - <a href="#about">About</a>
        - <a href="#how-to">How To... (button)</a>
        - <a href="#report-issue">Report Issue (button)</a>

<a id="user-content-dependencies" class="anchor" href="#dependencies" aria-hidden="true"> </a>
## Dependencies
|Software                        |Dependency                 |
|:-------------------------------|:--------------------------|
|[Microsoft Visual Studio Community 2017](https://www.visualstudio.com/vs/whatsnew/)|Solution|
|[Microsoft Office Developer Tools](https://blogs.msdn.microsoft.com/visualstudio/2015/11/23/latest-microsoft-office-developer-tools-for-visual-studio-2015/)|Solution|
|[Microsoft Excel 2010 (or later)](https://www.microsoft.com/en-au/software-download/office)|Project|
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

<body>

<a id="user-content-functionality" class="anchor" href="#functionality" aria-hidden="true"> </a>
## Functionality
This Excel ribbon named “Server Actions” is inserted after the “Home” tab when Excel opens.  Listed below is the detailed functionality of this application and its components.  

<a id="user-content-ping-test" class="anchor" href="#ping-test" aria-hidden="true"> </a>
###	Ping Test (Group)

<a id="user-content-ping" class="anchor" href="#ping" aria-hidden="true"> </a>
####	Ping (Button)
* This will ping the visible servers in the active table.

<a id="user-content-server-column" class="anchor" href="#server-column" aria-hidden="true"> </a>
####	Server (Dropdown)
* A list of column names from the active table.

<a id="user-content-ping-column" class="anchor" href="#ping-column" aria-hidden="true"> </a>
####	Ping (Dropdown)
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

<a id="user-content-comment" class="anchor" href="#comment" aria-hidden="true"> </a>
####	Comment (Dropdown)
* A list of column names from the active table.

<a id="user-content-group" class="anchor" href="#group" aria-hidden="true"> </a>
####	Group (Dropdown)
* A list of column names from the active table. This is used to group the servers in the remote desktop mananger file.

<a id="user-content-file-name" class="anchor" href="#file-name" aria-hidden="true"> </a>
####	File Name (Textbox)
* The file name to save the list of servers for Remote Desktop Manager.

<a id="user-content-options" class="anchor" href="#options" aria-hidden="true"> </a>
###	Options (Group)

<a id="user-content-rebuild-list" class="anchor" href="#rebuild-list" aria-hidden="true"> </a>
####	Rebuild Server List (Button)
* Rebuilds the server list from a LDAP query stored in the settings

<a id="user-content-refresh-lists" class="anchor" href="#refresh-lists" aria-hidden="true"> </a>
####	Refresh Dropdowns (Button)
* Refreshes all the dropdown values from the active table column names.

<a id="user-content-settings" class="anchor" href="#settings" aria-hidden="true"> </a>
####	Add-In Settings (Button)
* Opens the settings form

<h1 align="left">
  <img src="Images/ReadMe/settings_form.png" alt="settings" />
</h1>

<a id="user-content-visual-basic" class="anchor" href="#visual-basic" aria-hidden="true"> </a>
####	Visual Basic (Button)
* Opens the Visual Basic editor.

<a id="user-content-about" class="anchor" href="#about" aria-hidden="true"> </a>
###	About (Group)

<a id="user-content-how-to" class="anchor" href="#how-to" aria-hidden="true"> </a>
####	How To... (Button)
* Opens the how to guide in a browser

<a id="user-content-report-issue" class="anchor" href="#report-issue" aria-hidden="true"> </a>
####	Report Issue (Button)
* Opens the new issue page in a browser

#### Description (Label)
* The application name with the version

#### Release Date (Label)
* The release date of the application

#### Copyright (Label)
* The author’s name

</body>
</html>
