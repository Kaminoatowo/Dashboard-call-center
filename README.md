#Call center dashboard in Excel

This project is meant to learn how to use Excel for a simple data analysis project

##The files

###Call Center.csv

The **Call Center.csv** file contains about $33000$ records of calls. 

Each call is stored in a row of the *csv* file and each of these rows has $12$ columns.

Calls have an identifier string `id`, the name of the customer `customer_name`, the perceived sentiment during the call `sentiment`, an integer score value `csat_score`, the date the call happened `call_timestamp`, the reason of the call `reason`, the city of the customer `city`, the state where the city is `state`, the contact channel used `channel`, the response time `response_time`, the duration of the call in minutes `call duration in minutes` and the place of the call center `call_center`

##Procedure

###Observe data and make sure all is good

**Let's standardize entries in the various columns** (*if uploading the csv in excel through 'Data > from text/csv' a standard format is automatically imposed*)

1. the `call_timestamp` column has dates either aligned to the left or to the right: let's align all of them to the right; 

2. next, the same column has a **general** data type: let's make it a **date** data type

3. now I'll set numbers in the **number** data type (`csat_score`, `call duration in minutes`); I also removed the decimals (only zero) by specifying it in the Number tab

4. all of the data were collected in October 2020, so no need to use this redundant information: let's create a `Call_Day` column that contains only the day of the call

###Pivot tables and build charts

Let's build charts and Pivot Tables that will later be used in the dashboard. I'll create $5$ charts:

-   a line chart for the call trends
-   a doughnut chart for the call channel
-   a map for the states
-   a column chart for the reason of calling
-   a bar chart for the response time

(I may do something with sentiment, like an appreciation bar, or some emojis with transparency depending on the amount of sentiment expressed)

####Line chart

Show the trends of calls by listing the count of call for each day

1. Using the Pivot Table option, I create a table with rows labelled by `Call_Day` and the count of id as values

2. While having the Pivot Table selected (or a cell belonging to it) I can access the 'Pivot Table analysis' menu; look for 'Pivot chart' and create a line chart out of data stored in the pivot table

####Doughnut chart

Show the preferred communication channels 

1. Create a Pivot table with `id` values and `channel` as rows

2. Make a doughnut chart out of those data

####Map

Show the frequency of calls for each state

1. Create a Pivot table with `state` as rows and `id` as values

2. I can't create a map from the pivot table as I did before!! I need to copy the data in the pivot table somewhere else, which also means in the neighbouring cells!

3. Click on any of the cell of the new table and then "Insert > Map > Filled map"; data will be sent to bing! and will be elaborated and a map with number of calls will be created

4. Data that have been used to create the map aren't connected to the pivot table; let's connect them now: click on the map, go to "Chart Design > Select Data" and select data from the pivot table

####Reason of calling

Show a column bars chart with the reason of calling

1. Create a Pivot table with `reason` as rows and `id` as values

2. Make a column chart with the data

####Response time

Show the response time 

1. Create a Pivot table with `response_time` as rows and `id` as values

2. Make a bar chart with the data

##Other info

###How to use regular expressions in Excel

1. allow 'Developing' menu in excel by choosing the option in the 'File > Navigation bar' menu

2. select 'Visual Basic' in the 'Developing' menu

3. copy the 'Module' file in the [excel-regex-examples.xlsm] file to the excel project

4. the project has now $4$ new functions: *RegExpMatch*, *RegExpExtract* and *RegExpReplace*; they use regular expressions syntax to match, extract or replace expressions