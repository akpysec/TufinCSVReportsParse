# TufinCSVReportsParse

Automate Security Audits for Tufin Reports.
Put all your exported Reports.csv in the same folder specify path within the script and execute the script.
The script will produce .xslx file containing multiple sheets.

* May be written in more elegant fashion,
* More logic can be added

Checks are spread across the sheets:

- Any Service
  Logic: if "any" in service field
         if rule in enabled state
         if action is allow / accept
         if service is not negated
         if application identity cell is empty / any
         
- Any Source
  Logic: if "any" in source field
         if source user cell is empty / any
         if rule in enabled state
         if action is allow / accept
         if source is not negated
         if from zone cell is empty / any
         
- Any Destination
  Logic: if "any" in destination field
         if rule in enabled state
         if action is allow / accept
         if destination is not negated
         if to zone cell is empty / any
         
- Disabled Rules
  Logic: if rule in disabled state
  
- Rules with Reject option
  Logic: if rule in enabled state
         if action is reject
 
- Rules with No Logs
  Logic: if rule in enabled state
         if track cell contains none value
         
- Crossed rules
  Logic: if source zone in destination zone field
         if source in destination field
         if rule in enabled state
         if action is allow / accept
         if rules have the same service
         
- Un-Safe protocols (You may add to list as you wish)
  Logic: if service field contains services difined in a list
         if rule in enabled state
         if action is allow / accept
         
- Worst rules
         if "any" object is in source, destination & service fields
         if rule in enabled state
         if action is allow / accept
         
Enjoy :)
