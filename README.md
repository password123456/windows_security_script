# Windows Security Check Scripts

Overview:
- Windows security check scripts.
- Within the script the UAC execution codes included. so you don't have to worry about UAC Execution. just run and click confirm button when the UAC msgbox appear.
- I hope this helps to keep your windows system security's.

License:
 - Freeware, but You always need to cite your sources.

CheckList:
- You can check most of the recommendations of the Security Guide.
- See the list below.

```sh
1. Password Policy
  - maximunm password age.
  - minimum password age.
  - minimum password length.
  - password history count.
  - password complexity.

2. Account Lockout Policy
  - account lockout count.
  - account lockout duration.
  - reset account lockout count.

3. System default accounts status
  - default Administrator account name.
  - default guest account name.
  - enable/disable default adiministrator account.
  - enable/disable default guest account.

4. Event Audit Policy
  - enable/disable event audit items.

5. Remote Desktop
  - remote desktop status, port number

6. Default share 
  - status of default share 
   ex) admin$, IPC$

6. Local Administrators Group Account lists
  - count,list of Administrators rights account 

7. Stats of Unnecessary Windows Services
  - unnecessary windows service lists.
  - running status of it. 

7. various of system security setting
  - windows event log set size.
  - screen warning messages status when logon.
  - screen saver setting.
  - shutdownwithoutlogon
  - tasks schduler logging status
  - password expiry warning 
  - etc..

8. Widnows security updates status
  - scan the status of  the security updates using wsusscn2.cab
```
[![Hits](https://hits.seeyoufarm.com/api/count/incr/badge.svg?url=https%3A%2F%2Fgithub.com%2Fpassword123456%2Fhit-counter&count_bg=%2379C83D&title_bg=%23555555&icon=&icon_color=%23E7E7E7&title=hits&edge_flat=false)](https://hits.seeyoufarm.com)

