v1.1 - Box XML Generation Bug Fix
- Fixed Box XML generation not proprerly creating <Name> tag on itemparam xml rows

v1.2 - New MyShop Lib & DB Gen Tool
- Added new libcmgds_e standalone generation tool for set and individual items as well as SQL generation
- Fixed libcmgds_e generation to be properly created for each listing
- Fixed all libscmgds_e listing as invisible to start when generated
- Fixed Row Update Counter tool not respecting libcmgds_e
- Fixed PresentItemGenerator not respecting preferences

v1.3 - Mass Variable Manip Overhaul
- Fixed Mass Variable Manip bugs
- Allowed user in Mass Variable Manip to use to use blank ("Set to Fixed Value" for tehe source and "Math 
	Expression" for the destination) and wildcard (* for Regex Replace)
- Replaced the Mass Variable Manip UI with something more function-dependent that informs the user
	instead of the generic "To" and "From" fields
- Added CSV Import to Mass Variable Manip for additional conditional restrictions (e.g. by ItemID)
- Conditional statements in Mass Variable Manip now are manually toggleable (just in case) and set disabled
	by default unless you change the first one, which only that one will be toggled on.
- Changed default conditional statements to be just one and allow the user to add or remove more
- Auto-set conditional statements to OFF by default on the first one ONLY and then if you modify it in any
	way, it will set ON. When you add secondary statements, they are set ON by default and turn back ON when
	you change any of their values.

v1.4 - Hotfix
- Fixed rows having random new lines added to them in Mass Variable Manip