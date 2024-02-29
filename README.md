Here we solve the network flows optimization problem propused by AIMMS in https://download.aimms.com/aimms/download/manuals/AIMMS3_OM.pdf using Solver in Excel.

In this problem, we have two sources, located in Arnhem and Gouda, and six customers, located in London, Berlin, Maastricht, Amsterdam, Utrecht and The Hague. For reasons of efficiency, deliveries abroad are only made by one source. So that, Arnhem delivers to Berlin and Gouda to London.

We can find the original limits of production/supply for each source, the demand for each customer and transportation costs between each source and customer in solve_data_0.xlsm, input data tables. In this problem, the goal is to satisfy the customers’ demand while minimizing transportation costs.

You can review the problem formulation in problem_formulation_1.PNG and problem_formulation_2.PNG:
- The decision variables are in K10:P11.
- The objective function is located in L4 (this cell has a formula) and involves K10:P11 (decision variables) and C24:H25 (input data).
- The production limit for each source constraints involve Q10:Q11 (these cells have a formula with a sum for some decision variables) and C7:C8 (input data).
- The demand limit for each customer constraints involve K12:P12 (these cells have a formula with a sum for some decision variables) and C13:C18 (input data).
- If there is a minimum quantity that is mandatory to move between each source and each customer, we can define it in C32:H33, and decision variables (K10:P11) will have to satisfy it.

We have four files with different data. However, we could use one of these files, edit the input data and run the optimization using the blue button in order to get new results.
* solve_data_0.xlsm. This is the base case and we have selected information about sensibility analysis when we solved it (see sensibility_analysis_selection.PNG).
* solve_data_1.xlsm. Using base case, we move one ton of supply capacity from Gouda to Arnhem and the objective function improves in 0.2 euros (shadow price for Arnhem in the base case).
* solve_data_2.xlsm. Using the base case, we increase the demand in London in one ton and we increase one ton of supply capacity in Gouda (Gouda is the only source for London). Then the objective function gets worse in 2.5 euros (shadow price for London in the base case).
* solve_data_3.xlsm. Using the base case, we increase the demand in Berlin in one ton and we increase one ton of supply capacity in Gouda (Arnhem is the only source for Berlin). Then the objective function gets worse in 2.7 euros (shadow price for Berlin in the base case).
* solve_data_4.xlsm. Using the base case, we fixed a minimum transportation between Arnhem and Amsterdam equal to 1 ton. Then the objective function gets worse in 0.6 euros (reduced cost for this transportation in the base case).

To run the optimization:
- You have to install Solver complement in Excel.
- You have to allow running macros in Excel.
- And Solver has to be activated as a Reference for VBAProject.