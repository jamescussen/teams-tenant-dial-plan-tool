Teams Tenant Dial Plan Tool
===========================

            

The Tenant Dial Plan Tool is a PowerShell based tool that allows you to configure and edit Tenant Dial Plans within Office 365 for use with Microsoft Teams Direct Routing and Calling Plans. This tool is a sister tool to my
[Microsoft Teams Direct Routing Tool](https://www.myskypelab.com/2019/02/microsoft-teams-direct-routing-tool.html) that allows you to configure all the routing for Direct Routing within Office 365.

 

![Image](https://github.com/jamescussen/teams-tenant-dial-plan-tool/raw/master/TeamsTenantDialPlanEditorv1.00-400px.png)


**Version 1.01**
  *  The Skype for Business PowerShell module is being deprecated and the Teams Module is finally good enough to use with this tool. As a result, this tool has now been updated for use with the Microsoft Teams PowerShell Module version 2.3.1 or above.



**Tool Features**


  *  Log into O365 using the Connect SfBO button in the top left of the tool. Note: the Skype for Business Online PowerShell module needs to be installed on the PC that you are connecting from. You can get the module from here: https://www.microsoft.com/en-us/download/details.aspx?id=39366

  *  Create/Edit and Remove Tenant Dial Plan policies using the New.., Edit.. and Remove buttons.

  *  Copy existing Tenant Dial Plans and all their Normalisation rules to a new Tenant Dial Plan.

  *  Add/Edit Tenant Dial Plan normalisation rules. If the rule you are setting has a name that matches an existing rule, then the existing rule will be edited. If the rule’s name does not match an existing rule then it will be added as a new rule to the
 list. 
  *  Delete one or all normalisation rules from a Tenant Dial Plan policy. 
  *  Easily change the priority of normalisation rules with the UP and DOWN buttons.

  *  Test the normalisation rules! Teams currently (at the time of writing this) doesn’t have any normalisation rule testing capabilities. So I wrote a custom testing engine into the tool providing this feature. By entering a number into the Test textbox
 and pressing the Test Number button, the tool will highlight all of the rules that match in the currently selected Dial Plan that match in blue. The rule that has the highest priority and matches the tested number will be highlighted in green. The pattern
 and translation of the highest priority match (the one highlighted in green) will be used to do the translation on the Test Number and the resultant translated number will be displayed in the Test Result.


**For further details on the tool visit the blog post: [https://www.myteamslab.com/2019/09/teams-tenant-dial-plan-tool.html](https://www.myskypelab.com/2019/09/teams-tenant-dial-plan-tool.html)**


 






        
    
