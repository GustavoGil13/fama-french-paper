cls
clear

capture log close

set linesize 90
set scheme s2mono

graph drop _all
set graphics off

// alter the path to where the excel file is
local path = "C:\Users\Legion\Desktop"

if "`path'" == "" {
	di "Alter the path variable to the directory where the excel is"
	exit
}

cd "`path'"

// if save = 1 saves all matrix´s
local save = 1

// Data importing
import excel "database.xlsx", firstrow clear

// formating the time
tostring date, generate(dates)
replace date = date(dates, "YM")
format %td date
drop dates

tsset time

// -----------------------------------------------------------------------------
// -----------------------Fama and French Paper reprodution --------------------
// ---------------------- 7 July 1963 through December 1991 --------------------
// -----------------------------------------------------------------------------

// if dates are between the time interval place a 1
generate select_reprodution = .
replace select_reprodution = 1 if inrange(date, mdy(7,1,1963), mdy(12,1,1991))

// ----------------------------------- TABLE 2 ---------------------------------

// market returns
gen rm = rmrf + rf

// mean, standard-deviation, t test, autocorrelation lags 1, 2 e 12
foreach var of varlist rm rmrf hml smb {
	di `var'

	quietly: tabstat `var' if select_reprodution == 1, stat(mean sd) save
	local `var'_med = r(StatTotal)[1, 1]
	local `var'_std = r(StatTotal)[2, 1]

	quietly: ttest `var' = 0 if select_reprodution == 1
	local `var'_t_statistic = r(t)

	quietly: ac `var' if select_reprodution == 1, lags(12) generate(lag_value_`var')
	local lag_value_1_`var' =  lag_value_`var'[1]
	local lag_value_2_`var' = lag_value_`var'[2]
	local lag_value_12_`var' = lag_value_`var'[12]
	
	if `var' == rm {
		matrix table_1 = (``var'_med',``var'_std',``var'_t_statistic',`lag_value_1_`var'',`lag_value_2_`var'',`lag_value_12_`var'')
		}
	else {
		matrix table_1 = table_1 \ (``var'_med',``var'_std',``var'_t_statistic',`lag_value_1_`var'',`lag_value_2_`var'',`lag_value_12_`var'')
		}
		
}
matrix colnames table_1 = "Mean" "Standard-Deviation" "t test" "Lag 1" "Lag 2" "Lag 12"
matrix rownames table_1 = "RM" "RM-RF" "HML" "SMB"

// correlation matrix
quietly: corr rmrf smb hml if select_reprodution == 1
matrix correlations = r(C)
matrix rownames correlations = "RM-RF" "SMB" "HML"
matrix colnames correlations = "RM-RF" "SMB" "HML"


// generate excessive returns
foreach var in me1bm1 me2bm1 me3bm1 me4bm1 me5bm1 me1bm2 me2bm2 me3bm2 me4bm2 me5bm2 me1bm3 me2bm3 me3bm3 me4bm3 me5bm3 me1bm4 me2bm4 me3bm4 me4bm4 me5bm4 me1bm5 me2bm5 me3bm5 me4bm5 me5bm5{
gen excessive_return_`var' = `var' - rf
}

// mean, standard-deviation, t test
foreach var in me1bm1 me2bm1 me3bm1 me4bm1 me5bm1 me1bm2 me2bm2 me3bm2 me4bm2 me5bm2 me1bm3 me2bm3 me3bm3 me4bm3 me5bm3 me1bm4 me2bm4 me3bm4 me4bm4 me5bm4 me1bm5 me2bm5 me3bm5 me4bm5 me5bm5 {
	
	quietly: tabstat excessive_return_`var' if select_reprodution == 1, stat(mean sd) save
	local excessive_return_`var'_med = r(StatTotal)[1, 1]
	local excessive_return_`var'_std = r(StatTotal)[2, 1]
	
	quietly: ttest excessive_return_`var' = 0 if select_reprodution == 1
	local excessive_return_`var'_t = r(t)
	
	if `var' == me1bm1 {
		matrix table_2 = (`excessive_return_`var'_med',`excessive_return_`var'_std',`excessive_return_`var'_t')
		}
	else {
		matrix table_2 = table_2 \ (`excessive_return_`var'_med',`excessive_return_`var'_std',`excessive_return_`var'_t')
		}	
}

matrix colnames table_2 = "Mean" "Standard-Deviation" "t test"
matrix rownames table_2 = me1bm1 me2bm1 me3bm1 me4bm1 me5bm1 me1bm2 me2bm2 me3bm2 me4bm2 me5bm2 me1bm3 me2bm3 me3bm3 me4bm3 me5bm3 me1bm4 me2bm4 me3bm4 me4bm4 me5bm4 me1bm5 me2bm5 me3bm5 me4bm5 me5bm5


// ----------------------------------- TABLE 4 ---------------------------------

// Classic CAPM regression
foreach var in me1bm1 me2bm1 me3bm1 me4bm1 me5bm1 me1bm2 me2bm2 me3bm2 me4bm2 me5bm2 me1bm3 me2bm3 me3bm3 me4bm3 me5bm3 me1bm4 me2bm4 me3bm4 me4bm4 me5bm4 me1bm5 me2bm5 me3bm5 me4bm5 me5bm5 {

	quietly: reg excessive_return_`var' rmrf if select_reprodution == 1
	
	local b = r(table)[1,1]
	local t_b = r(table)[3,1]
	local r_2 = e(r2)
	local rmse = e(rmse)
	
	if `var' == me1bm1 {
		matrix table_4 = (`b',`t_b',`r_2',`rmse')
		}
	else {
		matrix table_4 = table_4 \ (`b',`t_b',`r_2',`rmse')
		}		
}

matrix colnames table_4 = "b" "t(b)" "R_2" "s(e)"
matrix rownames table_4 = me1bm1 me2bm1 me3bm1 me4bm1 me5bm1 me1bm2 me2bm2 me3bm2 me4bm2 me5bm2 me1bm3 me2bm3 me3bm3 me4bm3 me5bm3 me1bm4 me2bm4 me3bm4 me4bm4 me5bm4 me1bm5 me2bm5 me3bm5 me4bm5 me5bm5


// ----------------------------------- TABLE 6 ---------------------------------

// Three factor regression
foreach var in me1bm1 me2bm1 me3bm1 me4bm1 me5bm1 me1bm2 me2bm2 me3bm2 me4bm2 me5bm2 me1bm3 me2bm3 me3bm3 me4bm3 me5bm3 me1bm4 me2bm4 me3bm4 me4bm4 me5bm4 me1bm5 me2bm5 me3bm5 me4bm5 me5bm5 {

	quietly: reg excessive_return_`var' rmrf smb hml if select_reprodution == 1
	
	local b = r(table)[1,1]
	local t_b = r(table)[3,1]
	local s = r(table)[1,2]
	local t_s = r(table)[3,2]
	local h = r(table)[1,3]
	local t_h = r(table)[3,3]
	local r_2 = e(r2)
	local rmse = e(rmse)
	
	if `var' == me1bm1 {
		matrix table_6 = (`b',`t_b',`s',`t_s',`h',`t_h',`r_2',`rmse')
		}
	else {
		matrix table_6 = table_6 \ (`b',`t_b',`s',`t_s',`h',`t_h',`r_2',`rmse')
		}		
}

matrix colnames table_6 = "b" "t(b)" "s" "t(s)" "h" "t(h)" "R_2" "s(e)"
matrix rownames table_6 = me1bm1 me2bm1 me3bm1 me4bm1 me5bm1 me1bm2 me2bm2 me3bm2 me4bm2 me5bm2 me1bm3 me2bm3 me3bm3 me4bm3 me5bm3 me1bm4 me2bm4 me3bm4 me4bm4 me5bm4 me1bm5 me2bm5 me3bm5 me4bm5 me5bm5

// -----------------------------------------------------------------------------
// -----------------------Fama and French Paper extension ----------------------
// ---------------------- 1 Jan 1992 through Octouber 2016 ---------------------
// -----------------------------------------------------------------------------

// if dates are between the time interval place a 1
generate select_extension = .
replace select_extension = 1 if inrange(date, mdy(1,1,1992), mdy(10,1,2016))

// ----------------------------------- TABLE 2 ---------------------------------

// mean, standard-deviation, t test, autocorrelation lags 1, 2 e 12
foreach var of varlist rm rmrf hml smb {
	di `var'

	quietly: tabstat `var' if select_extension == 1, stat(mean sd) save
	local `var'_med = r(StatTotal)[1, 1]
	local `var'_std = r(StatTotal)[2, 1]

	quietly: ttest `var' = 0 if select_extension == 1
	local `var'_t_statistic = r(t)

	drop lag_value_`var'
	quietly: ac `var' if select_extension == 1, lags(12) generate(lag_value_`var')
	local lag_value_1_`var' =  lag_value_`var'[1]
	local lag_value_2_`var' = lag_value_`var'[2]
	local lag_value_12_`var' = lag_value_`var'[12]
	
	if `var' == rm {
		matrix extension_table_1 = (``var'_med',``var'_std',``var'_t_statistic',`lag_value_1_`var'',`lag_value_2_`var'',`lag_value_12_`var'')
		}
	else {
		matrix extension_table_1 = extension_table_1 \ (``var'_med',``var'_std',``var'_t_statistic',`lag_value_1_`var'',`lag_value_2_`var'',`lag_value_12_`var'')
		}
		
}
matrix colnames extension_table_1 = "Mean" "Standard-Deviation" "t test" "Lag 1" "Lag 2" "Lag 12"
matrix rownames extension_table_1 = "RM" "RM-RF" "HML" "SMB"

// correlation matrix
quietly: corr rmrf smb hml
matrix extension_correlations = r(C)
matrix rownames extension_correlations = "RM-RF" "SMB" "HML"
matrix colnames extension_correlations = "RM-RF" "SMB" "HML"

// mean, standard-deviation, t test
foreach var in me1bm1 me2bm1 me3bm1 me4bm1 me5bm1 me1bm2 me2bm2 me3bm2 me4bm2 me5bm2 me1bm3 me2bm3 me3bm3 me4bm3 me5bm3 me1bm4 me2bm4 me3bm4 me4bm4 me5bm4 me1bm5 me2bm5 me3bm5 me4bm5 me5bm5 {
	
	quietly: tabstat excessive_return_`var' if select_extension == 1, stat(mean sd) save
	local excessive_return_`var'_med = r(StatTotal)[1, 1]
	local excessive_return_`var'_std = r(StatTotal)[2, 1]

	quietly: ttest excessive_return_`var' = 0 if select_extension == 1
	local excessive_return_`var'_t = r(t)
	
	if `var' == me1bm1 {
		matrix extension_table_2 = (`excessive_return_`var'_med',`excessive_return_`var'_std',`excessive_return_`var'_t')
		}
	else {
		matrix extension_table_2 = extension_table_2 \ (`excessive_return_`var'_med',`excessive_return_`var'_std',`excessive_return_`var'_t')
		}	
}

matrix colnames extension_table_2 = "Mean" "Standard-Deviation" "t test"
matrix rownames extension_table_2 = me1bm1 me2bm1 me3bm1 me4bm1 me5bm1 me1bm2 me2bm2 me3bm2 me4bm2 me5bm2 me1bm3 me2bm3 me3bm3 me4bm3 me5bm3 me1bm4 me2bm4 me3bm4 me4bm4 me5bm4 me1bm5 me2bm5 me3bm5 me4bm5 me5bm5

// ----------------------------------- TABLE 4 ---------------------------------

// Classic CAPM regression
foreach var in me1bm1 me2bm1 me3bm1 me4bm1 me5bm1 me1bm2 me2bm2 me3bm2 me4bm2 me5bm2 me1bm3 me2bm3 me3bm3 me4bm3 me5bm3 me1bm4 me2bm4 me3bm4 me4bm4 me5bm4 me1bm5 me2bm5 me3bm5 me4bm5 me5bm5 {

	quietly: reg excessive_return_`var' rmrf if select_extension == 1
	
	local b = r(table)[1,1]
	local t_b = r(table)[3,1]
	local r_2 = e(r2)
	local rmse = e(rmse)
	
	if `var' == me1bm1 {
		matrix extension_table_4 = (`b',`t_b',`r_2',`rmse')
		}
	else {
		matrix extension_table_4 = extension_table_4 \ (`b',`t_b',`r_2',`rmse')
		}		
}

matrix colnames extension_table_4 = "b" "t(b)" "R_2" "s(e)"
matrix rownames extension_table_4 = me1bm1 me2bm1 me3bm1 me4bm1 me5bm1 me1bm2 me2bm2 me3bm2 me4bm2 me5bm2 me1bm3 me2bm3 me3bm3 me4bm3 me5bm3 me1bm4 me2bm4 me3bm4 me4bm4 me5bm4 me1bm5 me2bm5 me3bm5 me4bm5 me5bm5

// ----------------------------------- TABLE 6 ---------------------------------

// Three factor regression
foreach var in me1bm1 me2bm1 me3bm1 me4bm1 me5bm1 me1bm2 me2bm2 me3bm2 me4bm2 me5bm2 me1bm3 me2bm3 me3bm3 me4bm3 me5bm3 me1bm4 me2bm4 me3bm4 me4bm4 me5bm4 me1bm5 me2bm5 me3bm5 me4bm5 me5bm5 {

	quietly: reg excessive_return_`var' rmrf smb hml if select_extension == 1
	
	local b = r(table)[1,1]
	local t_b = r(table)[3,1]
	local s = r(table)[1,2]
	local t_s = r(table)[3,2]
	local h = r(table)[1,3]
	local t_h = r(table)[3,3]
	local r_2 = e(r2)
	local rmse = e(rmse)
	
	if `var' == me1bm1 {
		matrix extension_table_6 = (`b',`t_b',`s',`t_s',`h',`t_h',`r_2',`rmse')
		}
	else {
		matrix extension_table_6 = extension_table_6 \ (`b',`t_b',`s',`t_s',`h',`t_h',`r_2',`rmse')
		}		
}

matrix colnames extension_table_6 = "b" "t(b)" "s" "t(s)" "h" "t(h)" "R_2" "s(e)"
matrix rownames extension_table_6 = me1bm1 me2bm1 me3bm1 me4bm1 me5bm1 me1bm2 me2bm2 me3bm2 me4bm2 me5bm2 me1bm3 me2bm3 me3bm3 me4bm3 me5bm3 me1bm4 me2bm4 me3bm4 me4bm4 me5bm4 me1bm5 me2bm5 me3bm5 me4bm5 me5bm5

// ----------------------------------- save ------------------------------------

if `save' == 1 {
	
	capture mkdir excels

	putexcel set "excels/correlations.xlsx", replace
	putexcel A1 = matrix(correlations), names
	putexcel set "excels/table_1.xlsx", replace
	putexcel A1 = matrix(table_1), names
	putexcel set "excels/table_2.xlsx", replace
	putexcel A1 = matrix(table_2), names
	putexcel set "excels/table_4.xlsx", replace
	putexcel A1 = matrix(table_4), names
	putexcel set "excels/table_6.xlsx", replace
	putexcel A1 = matrix(table_6), names
	
	putexcel set "excels/extension_Correlations.xlsx", replace
	putexcel A1 = matrix(extension_correlations), names
	putexcel set "excels/extension_table_1.xlsx", replace
	putexcel A1 = matrix(extension_table_1), names
	putexcel set "excels/extension_table_2.xlsx", replace
	putexcel A1 = matrix(extension_table_2), names
	putexcel set "excels/extension_table_4.xlsx", replace
	putexcel A1 = matrix(extension_table_4), names
	putexcel set "excels/extension_table_6.xlsx", replace
	putexcel A1 = matrix(extension_table_6), names
	
}

matlist table_1
matlist correlations
matlist table_2
matlist table_4
matlist table_6

matlist extension_table_1
matlist extension_correlations
matlist extension_table_2
matlist extension_table_4
matlist extension_table_6