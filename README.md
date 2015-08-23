# Windows Security Check Scripts

Overview:
- windows security check scripts.
- Within the script the UAC execution codes included. so you don't have to worry about UAC Execution. just run and click confirm button when the UAC msgbox appear.
- i hope this helps to keep your windows system security's.

License:
 - Freeware, but You always need to cite your sources.

CheckList:
- You can check most of the recommendations of the Security Guide.
- see the list below.

```sh
1. Password Policy
  - maximunm password age.
  - minimum password age.
  - minimum password length.
  - password history count.
  - password complexity.

2. Account Lockout Policy
  - lockout count.
  - lockout duration.
  - reset lockout count.

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
