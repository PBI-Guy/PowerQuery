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

