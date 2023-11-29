# About me & the solution
This solution is a fork of existing solution (see section Credits) and has been developed in response to the need of having cool looking tabs with possibility of hiding tabs depending on the SharePoint user's group belonging. 
I am not a developer, just an enthusiast, so the solution is far from being optimal/well-written. It serves the purpose as of now :)
If you have some time, feel free to download it, clean it for yourself, redesign etc! 

# Prerequisites
The solution is based on the SPFX v1.18

# Features
Adding this webpart to the page allows to organize all different webparts in the same section into tabs. 
You can place many WebParts in single tab as well. 
Tabs can be configured - names can be adjusted as well as you can set the visiblity of the tab for specific SharePoint group members! (in a dirty way, too lazy to fix that :|)
In the edit mode you can find that each webpart on the page will have additional marking (WebPart #1, WebPart #2 etc.) - for better orientation which webpart to add to which tab.

![ModernTabs](https://github.com/aleksgabrysiak/spfx_modern_tabs_webpart/blob/main/ModernTabs.png)

![ModernTabs](https://github.com/aleksgabrysiak/spfx_modern_tabs_webpart/blob/main/ModernTabs.gif)

# Limitations:
Currently it is not possible to add the Modern Tabs webpart multiple times in the same section and embed it into another Modern Tab webpart - this is not supported, cannot predict what's gonna happen :)


# Credits
Original Solution: https://github.com/mrackley/Modern_Hillbilly_Tabs \AddTab script  CSS refered in the code: https://www.jqueryscript.net/other/Minimal-Handy-jQuery-Tabs-Plugin-AddTabs.html by dustinpoissant
# Disclaimer
**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**


# How to use the repo: 
- Clone this repository
- Ensure that you are at the solution folder
- in the command-line run:
  - **npm install**
  - **gulp serve**
