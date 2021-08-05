# styling-RAIS-Long-outputs

## To make analysis on longitudinal tables more easily, a small app has been developed to highlight cells of longitudinal tables based on indicators’ flags.
Though this app is developed through python, we managed to turn it into a standalone app, so that its’ runnable without having python installed.

So basically what this app does is that it process the longitudinal tables(.xlsx), and return a processed excel file with styles as follows:
-	If the cohort size of a series(row) is not Null and higher than user-defined limit, then the cell “year” will be filled by red.
-	If a cell is preliminary, then its font will be bold.
-	If a cell’s status is 6(data quality: acceptable), then its font-color will be red. 
-	If a cell’s status is 7(data quality: cautious), then its font-color will be deep red. 
-	If a cell’s release flag =1, then this cell will be filled by yellow.
Note: if a cell is preliminary and its status is 6/7 in the meantime, then its font will be bold + red/deep red.

## To run this app:
1.	Clicking here and wait for a moment, it may take a while to start the program.
2.	It shall pop up an interface and CMD (just ignore the CMD).
3.	In the interface, click “open file” to select an excel file (eg: pathway.xlsx) then you shall see the title of selected excel displace on the right panel(“tables to be processed”)
4.	Enter a number in the entry “limit of cohort size” (eg: 100), that’s the threshold used to determine if to highlight cell “year” of a given series (row).
5.	Click “process and save”, and it will ask you the path to save the output, after that you shall see the “status” change from “not started” to “start processing!”
6.	Once the “status” changed from “start processing!” to “task done!” ( It may take 110 seconds), you can found output file in the defined path.
