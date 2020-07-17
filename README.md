# DDLtoExcel
Creates an Excel workbook based on SQL DDL commands.

## Before You Proceed

### Background
I developed this system because we're working on a system with a large database and we don't have any documentation for it. The database used **Oracle**. It has multiple connected schemas - each schema has hundreds of tables and each table contains several columns. In order to understand how each table are connected to one another, I extracted the DDL commands via **sqldeveloper**. Then process each line of code to create an Excel workbook.

This helped me with two things:
1. Understanding the database structure.
2. Creating a documentation via an Excel file.

### Limitations
- It works only on Oracle database.
- It has limited [DDL command support](#supported-commands) (I only added those I needed at the moment).
- Each **Table Sheet** is named after the `TABLE` name, it has to be limited to **31 characters**, or else it will throw an error.
  - If the `SCHEMA` name is **INCLUDED** in the DDL commands, the name of **Table Sheet** will be *"SCH[n].TableName"* where *n* is the number of schema.
  - I decided not to include the `SCHEMA` name in the name of **Table Sheet** because of the **31-character** limitation of Excel.
- For multiple connected `SCHEMA`
  - If the `SCHEMA` name is **NOT INCLUDED** in the DDL commands, each `TABLE` **MUST** have a **UNIQUE** name, or else it will throw an error due to duplicate sheet names.

## Getting Started
These instructions will get you a copy of the project up and running on your local machine for development purposes.

### Prerequisites
- .NET Framework 4.7.2
- Microsoft Excel 2013
- Microsoft Visual Studio 2015

#### NOTE:
The above versions are what I used in developing the project.
- For **.NET Framework**, you can download the version I used or just change the settings on the project's *Properties*.
- For **Microsoft Excel**, you can change the *References* used in the project.
- For **Microsoft Visual Studio**, it needs to have the *.NET desktop development* workload installed.

### Opening and Executing the project
1. Clone this repository on your local machine.
2. Locate the project on your local machine, open *"DDLToExcel.sln"* via **Microsoft Visual Studio**.
3. **Build** the project, then **Start**.

## How It Works
The system creates the table structures based on the DDL commands on a .sql file.

### Supported Commands
- `CREATE TABLE`
- `CREATE GLOBAL TEMPORARY TABLE`
- `ALTER TABLE`
- `COMMENT ON TABLE`
- `COMMENT ON COLUMN`

It then creates an Excel workbook with:

#### Summary Sheet
  - It contains the list of `TABLE` and their corresponding `COMMENT`.
  - Each `TABLE` has a hyperlink to its corresponding **Table Sheet**.
  
*Example:*
|No.|Table Name|Comment|
|---|---|---|
|1|[Table1 Name](#table-sheet)|Table1 Comment|
|2|[Table2 Name](#table-sheet)|Table2 Comment|
|3|[Table3 Name](#table-sheet)|Table3 Comment|
|...|...|...|
|n|[Table[n] Name](/#)|Table[n] Comment|

#### Table Sheet
  - Each `TABLE` has it's own **Table Sheet**.
  - It contains the list of `COLUMN` and their corresponding `CONSTRAINT ` and `COMMENT`.
  - Each `FOREIGN KEY` constraint has a hyperlink to its corresponding **Table Sheet**.
  
*Example:*
<table>
  <thead>
    <tr>
      <th rowspan=2>No.</th>
      <th rowspan=2>Column Name</th>
      <th rowspan=2>Data Type</th>
      <th rowspan=2>Default</th>
      <th colspan=5>Constraint</th>
      <th rowspan=2>Comment</th>
    </tr>
    <tr>
      <th>Not Null</th>
      <th>Primary Key</th>
      <th>Unique</th>
      <th>Foreign Key</th>
      <th>Check</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td>1</td>
      <td>Column1 Name</td>
      <td>Column1 Data Type</td>
      <td>[AUTO INCREMENT] | [DEFAULT VALUE]</td>
      <td>[YES]</td>
      <td>[YES]</td>
      <td>[YES]</td>
      <td><a href="#table-sheet">[Table Name].[Column Name]</a></td>
      <td>[CONDITION]</td>
      <td>Column1 Comment</td>
    </tr>
    <tr>
      <td>2</td>
      <td>Column2 Name</td>
      <td>Column2 Data Type</td>
      <td>[AUTO INCREMENT] | [DEFAULT VALUE]</td>
      <td>[YES]</td>
      <td>[YES]</td>
      <td>[YES]</td>
      <td><a href="#table-sheet">[Table Name].[Column Name]</a></td>
      <td>[CONDITION]</td>
      <td>Column2 Comment</td>
    </tr>
    <tr>
      <td>3</td>
      <td>Column3 Name</td>
      <td>Column3 Data Type</td>
      <td>[AUTO INCREMENT] | [DEFAULT VALUE]</td>
      <td>[YES]</td>
      <td>[YES]</td>
      <td>[YES]</td>
      <td><a href="#table-sheet">[Table Name].[Column Name]</a></td>
      <td>[CONDITION]</td>
      <td>Column3 Comment</td>
    </tr>
    <tr>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
    </tr>
    <tr>
      <td>n</td>
      <td>Column[n] Name</td>
      <td>Column[n] Data Type</td>
      <td>[AUTO INCREMENT] | [DEFAULT VALUE]</td>
      <td>[YES]</td>
      <td>[YES]</td>
      <td>[YES]</td>
      <td><a href="#table-sheet">[Table Name].[Column Name]</a></td>
      <td>[CONDITION]</td>
      <td>Column[n] Comment</td>
    </tr>
  </tbody>
</table>
