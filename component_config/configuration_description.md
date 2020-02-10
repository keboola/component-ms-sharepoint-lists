### Warning: 
The writer currently supports only **`full load`**. 
All existing items in the destination list will be removed prior the load. 

### Creating new list

New list may be created directly from configuration. If the list already exists, the configuration section `Create new list`
 is ignored. Once created, changes made in that configuration section will ignored.
 
**NOTE**: When creating a new list the `genericList` template is used and the resulting list always contains a required column named `Title`

### Column definition
 
- Currently only `text`, `number`, `date`, `dateTime` column types are available when creating a new list.
- When no column parameters are specified, all columns will be created as `text` fields and the resulting column display names 
will match the input table.
- Title column name mapping is always required.
- Once created, any changes made to this section will be ignored. It is not possible to update column definition of an exisitng list.

### Writing to an existing list

When the `Create new list` section is empty, a list with the specified name is expected to exist, otherwise the job fails.

**NOTE**: During each execution all existing list items are removed from the destination list prior upload.


Additional documentation is available [here](https://bitbucket.org/kds_consulting_team/kds-team.wr-ms-sharepoint-lists/src/master/README.md).