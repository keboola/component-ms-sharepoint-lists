# Microsoft SharePoint Lists writer

Write and create SharePoint lists directly from Keboola Connection.

**Table of contents:**  
  
[TOC]

# Functionality notes

This component allows you to create a new SharePoint list directly from Keboola Connection or rewrite content of an existing one.

**Important Note**: The current version supports only `full load`, meaning that the contents of the existing list will always 
be pruned before upload.

## Creating new list

New list may be created directly from configuration. If the list already exists, the configuration section `Create new list`
 is ignored. Once created, changes made in that configuration section will ignored.
 
**NOTE**: When creating a new list the `genericList` template is used and the resulting list always contains a required column named `Title`

### Column definition
 
- Currently only `text`, `number`, `date`, `dateTime` column types are available when creating a new list.
- When no column parameters are specified, all columns will be created as `text` fields and the resulting column display names 
will match the input table.
- Title column name mapping is always required.
- Once created, any changes made to this section will be ignored. It is not possible to update column definition of an exisitng list.

## Writing to an existing list

When the `Create new list` section is empty, a list with the specified name is expected to exist, otherwise the job fails.

**NOTE**: During each execution all existing list items are removed from the destination list prior upload.

# Configuration
 
## Host base name

Your MS SharePoint host, typically something like `{my-tenant}.sharepoint.com`

## List name

Name of the SharePoint list you want to update, exactly as it is displayed in the UI (case sensitive). 

![List example](docs/imgs/list.png)

## List description

Optional list description.



## Create new list

### Title column

hen creating a new list the `genericList` template is used and the resulting list always contains a required column named `Title`.
For this reason it is required to define the name of the column in the input table that will be treated as the `Title` column. 
It may be for example you index column.

### List column parameters

Optional parameters of the list columns. If left empty, all fields will be treated as text. 
If specified list already exist, this section is ignored.

- **Source column name** - (REQ) Name of the column in the input table
- **Display name** - (REQ) Display name that will be visible in the SharePoint UI
- **Column data type**

#### Supported data types

- `text`
- `number`
- `dateTime`, `date` - ISO 8601 format is expected, however the default `YYYY-MM-DD` format should work as well. 
The value is then displayed in SharePoint with default formatting. 

### Include additional system columns

SharePoint lists also contain system columns that are not visible in the default UI view. By selecting this option it is possible to retrieve these 
columns also. **NOTE** By default, the extractor uses the `Display Name` of columns. It may happen that some of the system or custom columns, 
share the same display name. In such case, the extractor automatically deduplicates the column names by appending the underlying unique `API name` 
of the column separated by the `_` underscore sign. Leading to column names such as `Title_LinkTitleNoMenu`, `Title_LinkTitle`.

### Use column display names

List columns in SharePoint consists of a `display name` that is by default displayed in the SharePoint UI view 
and an `api name` that is underlying unique identifier of a column and does not change. This option 
allows you to choose which of the names you wish to use in the result table. 

**NOTE** that when using the `display name` duplicate column names will be automatically deduplicated. 
Also some of the system columns are prefixed with `_`, these underscores will be dropped since the *Storage* columns cannot
 start with underscore signs.

# Development
 
This example contains runnable container with simple unittest. For local testing it is useful to include `data` folder in the root
and use docker-compose commands to run the container or execute tests. 

If required, change local data folder (the `CUSTOM_FOLDER` placeholder) path to your custom path:
```yaml
    volumes:
      - ./:/code
      - ./CUSTOM_FOLDER:/data
```

Clone this repository, init the workspace and run the component with following command:

```
git clone https://bitbucket.org:kds_consulting_team/kds-team.ex-ms-sharepoint.git my-new-component
cd my-new-component
docker-compose build
docker-compose run --rm dev
```

Run the test suite and lint check using this command:

```
docker-compose run --rm test
```

# Integration

For information about deployment and integration with KBC, please refer to the [deployment section of developers documentation](https://developers.keboola.com/extend/component/deployment/) 