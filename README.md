These are some scripts I've written at work to automate some tidious tasks of my team. 
If you want to use them read them through and replace all file paths, names of servers etc since they are made nonsense.

Excel to AD groups.ps1 is a script that takes an excel file (see reference file) and creates Active Directory groups based on the data. Script is to be run on local computer remoting to server. 

Excel to MyAccess entilements.ps1 is a script that creates entilements to IDM system.

PC deletion script.ps1 deletes computer identities from Active Directory and SCCM. Must be run on the server hosting them.

TeamViewer idle license removal.ps1 Removes any licenses from idle computers managed by the organization.  
