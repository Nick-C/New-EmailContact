New-EmailContact
===================


This is a pretty basic script to that will connect to an on-prem (no O365 support), Exchange server and create one or more mail enabled contacts and then grant a user Send As permissions over it.


Notes
-------------
By default this script will use the naming format *ExternalName - Internal UserName* for the contact as this is what was required when creating this script, if you don't want this then edit lines 156 and 167 to edit the naming format used.