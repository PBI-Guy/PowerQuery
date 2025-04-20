# Power Query

This repo is about showing the power of Power Query and how to transform your data with it. We will first focus on Excel and in a later stage, add Power BI as well.

## Scenario Description

We're a global sales manager and want to analyze our sales performance across different dimensions like time, product, and many more. From our controlling team we receive a monthly CSV file with the current sales performance. With Power Query, we want to automate as much as possible to avoid manual tasks and be as efficient as possible. 

## Problem statement

As we receive each month one CSV file, it's hard to analyze revenue over time. Further, it takes a lot of manual effort as of now to combine the files to one table.

## Scenario 1 - Connect to one CSV file

With Power Query, we can connect to different data sources. In the first scenario, we connect to one CSV file and explore the possibilities. 

Before we begin, make sure to download the [Data](./Data/) folder and store it e.g. on your Desktop. 

Open Excel, select **Data** in the Ribbon, click on **Get Data**, choose **From File**, and hit **From Text/CSV**.

<img src="./PNG/01 Get Data and Connect to CSV.png" width="500">

A window will pop up in which you can select one CSV file. Navigate to the **CSV** folder and select **April 2018.csv**. Confirm by selecting **Import**.

<img src="./PNG/02 Select April CSV file.png" width="500">

Excel will now create a connection to the CSV file and show a preview of the data. In a first step, we will just **Load** the data as is.

<img src="./PNG/03 Load Data.png" width="500">

After hitting the load button, Excel will copy the data from the CSV file into Excel and create a new worksheet _April 2018_. Further, in the Queries & Connections pane, we will see one query _April 2018_ with 2249 rows loaded. In case you don't see the Queries & Connections pane, head over to **Data** and select **Queries & Connections**.

<img src="./PNG/04 Queries & Connections.png" width="500">

As we have now our data loaded into the Excel, we can easily create a Pivot Table. For that, click somewhere into your table (e.g. cell A2), select **Insert** from the Ribbon, and choose **PivotTable**.

<img src="./PNG/05 Create Pivot Table.png" width="500">

In the pop up window, we'll see the Table _April 2018_ is selected as well as to create the Pivot Table in a new worksheet. Confirm with **OK**.

<img src="./PNG/06 Configure Pivot Table.png" width="500">

Now, we can easily select the required fields. For example, let's choose **ProductCategory** and **SumOfAmount**. Make sure ProductCategory is added to the _Rows_ section and SumOfAmount is in the _Values_ section. This way, we get a quick overview which product category has the most and least sales. 

<img src="./PNG/07 Pivot Table.png" width="500">

If we now add **TranDate** to the _Columns_ section, we'll get a breakdown by date.

<img src="./PNG/08 TranDate added to Pivot Table.png" width="500">

As this is quite a wide range and chaotic, it would be better to have some kind of hierarchy - Year, Quarter, Month, Day. For that purpose, we can leverage Power Query to create the desired columns and avoid manual work in future. 
Let's launch Power Query by selecting **Data** from the Ribbon, click on **Get Data** and choose **Launch Power Query Editor...**.

<img src="./PNG/09 Launch Power Query.png" width="500">

A new window will pop up. This is the Power Query window in which we'll see on table already - our **April 2018** CSV file which we loaded previously. If you click on it, you'll get a preview of your data. 

<img src="./PNG/10 Power Query Editor.png" width="500">

One the left hand side (1), you'll see all table to which you have connected to. In the middle (2), you'll see a preview of your data. This is depending on which table you select on the left hand side. Lastly, at the right (3), you'll see all transformation steps applied to this specific table. As of now, we see three steps: Source - Promoted Headers - Changed Type. You can explore each step by selecting it and the preview will show you how the data looked like at this step. Once we refresh the data, Power Query will go through each step from top to down and perform each applied transformation. Once done, the data will be loaded into Excel. Due to that approach, we only need to modify and transform the data once and every time we refresh our data, all steps will be applied automatically. 

For our purpose, let's add a Year, Quarter, Month, and Date column. We start with the Year column by selecting the **TranDate** from our April 2018 Table, go to **Add Column** in the Ribbon, choose **Date**, click on **Year** and select **Year** from the options.

<img src="./PNG/11 Add Year column in Power Query.png" width="500">

This will add a new column **Year** as well as a new step into our _Applied Steps_ section. 

<img src="./PNG/12 New Year Column.png" width="500">

Repeat the steps for Quarter, Month, and Day. Keep in mind to select the TranDate before you add a new column.

Once finished, select **Home** in the Ribbon and hit **Close & Load**.

<img src="./PNG/13 Close and Load.png" width="500">

This way, Power Query will refresh our data with the new speps applied and add the four new columns.

<img src="./PNG/14 New Columns added.png" width="500">

Because the Pivot Table is connected to the April 2018 table, we need now to refresh the Pivot Table as well. To do so, go to your Pivot Table, select it, and press **Refresh** in the **PivotTable Analyze** Ribbon menu. Once the refresh has successfully completed, we'll see the four new columns. 

<img src="./PNG/15 Refresh Pivot Table.png" width="500">

Let's now adjust the Pivot Table to show Year and Month in the Columns section instead of days. As we can see, a hierarchy helps for a better overview. Obviously, we can also add Quarter and Day if needed. 

<img src="./PNG/16 Adjusted Pivot Table.png" width="500">

But as we wish to analyze our revenue across multiple months to compare and identify trends, we would now need to connect to the other CSV files. One approach would be to repeat the steps and combine the tables in Power Query or Excel. But as we want to be as efficient as possible, let's leverage the Power of Power Query even more.

## Scenario 2 - Connect to folder and combine multiple CSV files

In our existing file, we select **Get Data** from the **Data** menu in the Ribbon, click on **From File** and select **From Folder**. 

<img src="./PNG/17 Connect to Folder.png" width="500">

In the pop up window browse to the **CSV** folder and confirm with **Open**.

<img src="./PNG/18 Select CSV folder.png" width="500">

Now, we're connected to the folder. In the preview screen, we can see all files which are stored in this folder with some meta data like modified date, created date, path, etc. As we're not interested in the meta data but rather in the binaries, we have to extract those and combine it into one big table. Power Qurey helps us to achieve our goal by providing a **Combine** button. If we select it, it would do the magic automatically for us but in this case, we want to do further transformation steps on top so we select **Transform Data**. 

<img src="./PNG/19 Transform Data Folder Connector.png" width="500">

In Power Query, we see now two tables - one called _April 2018_ and the other one _CSV_. Select the **CSV** table and click on the two arrows in the **Content** column. 

<img src="./PNG/20 Combine Binaries in Power Query.png" width="500">

A new window will pop up in which we define our sample CSV. Based on it the schema will be defined (meaning which columns should be loaded). Per defalt, the first file is taken as sample but it can be overwritten at the top. In our case, we can leave it as is. 
Like before, when we connected to one single CSV file, we have now a preview of the data from the sample file. As the data looks good, we confirm with the **OK** button.

<img src="./PNG/21 Preview before combining Binaries.png" width="500">

After a few seconds, we have one big table containing all CSV data!

<img src="./PNG/22 Binaries combined.png" width="500">

> Power Query created automatically a function which is used to extract the binaries. As of now, we do not need to take care of. 

Let's do some further transformation before we load the data. First, we want to remove the _Source.Name_ column as it's not needed. **Right click** on the column name and select **Remove**.

<img src="./PNG/23 Remove Source.Name column.png" width="500">

Second, we want to rename some columns. **Double click** on _TranDate_ and rename it to **Date**. Please rename columns as follow:

| Original Name | Renamed |
|--------------|----------|
| TranDate | Date |
| Dept | Country |
| SumofAmount | Revenue |
| CustomerSegment | Customer Segment |
| ProductCategory | Product Category |

> All changes will not affect the CSV file and are only applied during the load of the data into Excel.

Once done, confirm by selecting **Close & Load** from the **Home** Ribbon.

<img src="./PNG/24 Close and Load.png" width="500">

A new worksheet with a new table _CSV_ has been loaded. Instead of creating now a new Pivot Table, let's leverage the existing one and just change the data source. For that, we go to our _Sheet2_ in which our Pivot Table is, click into the Pivot Table (e.g. cell A6), select **PivotTable Analyze** in the Ribbon, and click on **Change Data Source**. 

<img src="./PNG/25 Change Data Source in Pivot Table.png" width="500">

Change the Table/Range to **CSV** and confirm with **OK**. 

<img src="./PNG/26 Select CSV as table source.png" width="500">

Lastly, let's add **Product Category** to the _Rows_, **Revenue** to _Values_, and the automatic created field **Months (Date)** to the _Columns_ secion.

<img src="./PNG/27 Pivot Table on CSV files.png" width="500">

Now, we can easily compare the numbers for each month. Let's create a column chart to visualize our result. Select **Insert** in the Ribbon and click on **PivotChart**. 

<img src="./PNG/28 PivotChart.png" width="500">

In the pop up window select **Column**, make sure to select the first one, and confirm with **OK**.

<img src="./PNG/29 PivotChart Column Chart.png" width="500">

Next, right click in the chart and choose **Select Data**.

<img src="./PNG/30 Change Pivot Chart fields.png" width="500">

Switch now Rows and Columns with the button in the middle and confirm with **OK**. 

<img src="./PNG/31 Switch Row and Column in PivotChart.png" width="500">

This way the Month column is on the X-Axis. As the Pivot Chart and Pivot Table are now connected, we can filter our Pivot Table and influence our Chart this way. Let's test it by selecting the filter icon next to _Column Labels_ and filter down to **Decor** and **Furniture**.

<img src="./PNG/32 Set Filter on PivotTable.png" width="500">

Let's test the magic of Power Query and add a new CSV file into our folder. Go back to the **Data** folder and Drag & Drop the **September 2018.csv** file into the **CSV** folder.

<img src="./PNG/33 Save September 2018 into CSV folder.png" width="500">

Once done, go back to Excel and hit **Refresh All** from the **Data** Ribbon.

<img src="./PNG/34 Refresh Power Query and PivotTable.png" width="500">

If everything works as expected, September data will be visible in our Pivot Table as well as the Chart! This way, we can now just add new files into the CSV folder, hit the Refresh button in Excel and everything else will magically be done for us!

If we go one step further and change the Columns fields now to **Dept** and remove **Product Category**, we see the Table as well as the Chart is influenced. 

<img src="./PNG/35 Change Columns in Pivot Table.png" width="500">

For reporting purpose, it would be easier if we would have names instead of numbers for our departement respectively countries. There are different ways to achieve it. One would be to use Power Query and add an additional column with an IF statement. For example, if the country code is 110, write USA, if 120, write Germany, etc. The issue with this approach is if a new country code appears or changes, we need to adjust our Power Query. Therefore, let's use a better approach and build a data model.

## Scenario 3 - Create a data model

Within our CSV folder, we have a **Country Codes.csv** file which we can use in Excel and create a data model. We again connect to the Excel Sheet via **Get Data** form the **Data** Ribbon and select **From File** - **From Text/CSV**.

<img src="./PNG/01 Get Data and Connect to CSV.png" width="500">

On the next screen, browse to your **Data** folder and select **Country Codes.csv**.

<img src="./PNG/36 Connect to Country Codes.png" width="500">

On the next screen, just hit **Load**.

Afterwards, go to the **Data** Ribbon and select **Data Model** - **Relationships**.

<img src="./PNG/37 Open Relationship View.png" width="500">

In the _Manage Relationships_ window, select **New...*. 

<img src="./PNG/38 New Relationship.png" width="500">

Now, create a relationship between the two tables **CSV** and **Country_Codes** on the columns **Dept** and **Country Code** as shown in the screen shot below. Confirm with **OK**.

<img src="./PNG/39 Create relationship.png" width="500">

Confirm once more by selecting **Close**.

Excel offers the possibility of showing the relationship between the tables in a view. To access it, select **Data Model** from the **Data** Ribon and choose **Manage Data Model**. A new window pops up in which you can select the **Diagram View** to see a visual representation of the relationships between the tables. 

<img src="./PNG/40 Relationship Diagram View.png" width="500">

Let's close the Data Model window and return to our Sheet2. In here, we would like to add a new table - the Country Codes table - to our existing Pivot Table. To do so, choose **More Tables...** to create a new Pivot Table, which is required now.

<img src="./PNG/41 More Tables Pivot Table.png" width="500">

A message pops up which indicates we need to create a new Pivot Table. Confirm by selecting **Yes**. 

<img src="./PNG/42 Create new Pivot Table.png" width="500">

Now, all three tables are accessible. Add the **Country** column from the _Country Codes_ table to the _Columns_ section. Further, add the **Product Category** from the _CSV_ table to the _Rows_ section. Lastly, make sure **Revenue** is added to the _Values_. 

<img src="./PNG/43 New Pivot Table based on Data Model.png" width="500">

Lastly, we can add a Pivot Chart to visualize our data in a column chart. Just follow the same steps as before b selecting **Insert** from the Ribbon menu - **PivotChart** - **Column**. If we swap column and rows once more, we got a chart like the one in the picture below.

<img src="./PNG/44 PivotChart Final.png" width="500"> 

## Bonus - Add a Date Table as dimension

Following best practice and creating a start schema. it makes sense to add a Date table to our model instead of adding further columns like Year, Month, Day to our Fact (CSV) table. We can manually create one in Power Query. To do so, hit the **Get Data** button from the **Data** Ribbon and select **From Other Source** - **Blank Query**.

<img src="./PNG/45 Create Blank Query.png" width="500">

Power Query Editor window will appear. Here, select **Advanced Editor** from the **Home** Ribbon. Copy and paste the code below, confirm with **Done.**

<img src="./PNG/46 Advanced Editor Power Query.png" width="500">

```let
    Source = CSV,
    MinDate = List.Min(Source[Date]),
    MaxDate = List.Max(Source[Date]),
    StartDate = #date(Date.Year(MinDate), 1, 1),
    EndDate = #date(Date.Year(MaxDate), 12, 31),
    DateList = List.Dates(StartDate, Duration.Days(EndDate - StartDate) + 1, #duration(1, 0, 0, 0)),
    DateTable = Table.FromList(DateList, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    ChangedType = Table.TransformColumnTypes(DateTable,{{"Column1", type date}}),
    RenamedColumns = Table.RenameColumns(ChangedType,{{"Column1", "Date"}}),
    InsertYear = Table.AddColumn(RenamedColumns, "Year", each Date.Year([Date]), Int64.Type),
    InsertMonth = Table.AddColumn(InsertYear, "Month", each Date.Month([Date]), Int64.Type),
    InsertDay = Table.AddColumn(InsertMonth, "Day", each Date.Day([Date]), Int64.Type),
    InsertQuarter = Table.AddColumn(InsertDay, "Quarter", each Date.QuarterOfYear([Date]), Int64.Type),
    InsertWeekOfYear = Table.AddColumn(InsertQuarter, "WeekOfYear", each Date.WeekOfYear([Date]), Int64.Type)
in
    InsertWeekOfYear
```

The code uses the CSV table as source and extract the min and max year from the Date column. Based on it, it creates dynamically a date table from the first day of the min Year until the last day of the max Year. In our example this means it create automatically a date table from 01. January 2018 until 31. December 2018. 

Rename **Query1** to **Date** through the Properties on the right hand side (or double click on the table name).

<img src="./PNG/47 Rename Query1.png" width="500">

Now, hit the **Close & Apply** button from the **Home** Ribbon to load the new _Date_ table. 

Once loaded, create a relationship between the _Date_ and _CSV_ table by selecting **Data** in the Ribbon - **Data Model** - **Relationships** and hit **New...**. Make sure to create a new relationship between the _CSV_ and _Date_ table on the **Date** columns. Confirm by selecting **OK** and **Close** afterwards. 

<img src="./PNG/48 Create Relationship on Date table.png" width="500">

With our new dimension table **Date** and a correct relationship to our fact table **CSV** we can easily add any column to our Pivot Table and Chart. 

## Clean Up
As we don't need the _April 2018_ table anymore, we can easily delete it. To do so, open **Queries and Connections** through the **Data** Ribbon, right click on **April 2018**, and select **Delete**. 

<img src="./PNG/49 Delete April 2018 table.png" width="500"> 

Confirm by clicking on **Delete** on the pop up window. 

<img src="./PNG/50 Delete confirmation.png" width="500">

Delete the **April 2018** worksheet and **Sheet2** as well as it's not needed anymore. 