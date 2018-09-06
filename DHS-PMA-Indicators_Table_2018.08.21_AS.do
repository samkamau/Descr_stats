/*******************************************************************************
*
*  FILENAME:	DHS-PMA-Indicators_Table_date_initials.do
*  PURPOSE:		Generate the DHS, PMA indicators table, all rounds, all country
*  CREATED:		Aisha Siewe (asieweb1@jhu.edu)
*  DATA IN:		Pubicly released dataset / WWA when public dataset unavailable
*  DATA OUT:	CC_.KeyIndicators.xls		
*******************************************************************************/
capture clear all
cd "/Users/asiewe/Documents/Data Analysis/DHS-PMA-Indicators/Excel Output"
*local filedir "/Users/asiewe/Documents/Data Analysis/DHS-PMA-Indicators/Descr_stats"

/*
***********************************************************************************
*** 		BURKINA FASO
***********************************************************************************
capture clear all
local excel "BF_KeyIndicators.xlsx"
local excel_sheet "CPR, mCPR, unmet need"
putexcel set "`excel'", modify sheet("`excel_sheet'")

**** ROUND 1
clear
use "~/Dropbox (Gates Institute)/01_Burkina/PMABF_Datasets/Round1/Final_PublicRelease/HHQ/PMA2014_BFR1_HHQFQ_v3_10Aug2018/PMA2014_BFR1_HHQFQ_v3_10Aug2018.dta", clear
putexcel A8=("PMA2020")
putexcel B8=("Round 1")
putexcel C8=("11/2014- 01/2015")

** COUNT - Female Sample - All / Married Women  **
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel D8=matrix(FQresponse_1)
restore
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3) & (FQmarital_status==1 | FQmarital_status==2)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel Q8=matrix(FQresponse_1)
restore 

*** Estimate Percentage and 95% CI
keep if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
egen all=tag(FQmetainstanceID)
egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)

svyset EA_ID [pw=FQweight], strata(strata) singleunit(scaled)
foreach group in all mar {
preserve
	keep if `group'==1
	foreach indicator in cp mcp unmettot {
		svy: prop `indicator', citype(wilson)
		matrix reference=r(table)
		matrix `indicator'_`group'_percent=reference[1,2]*100	
		matrix `indicator'_`group'_se=reference[2,2]*100
		matrix `indicator'_`group'_ll=reference[5,2]*100
		matrix `indicator'_`group'_ul=reference[6,2]*100
		}	
	restore
	}
putexcel E8=matrix(cp_all_percent)
putexcel F8=matrix(cp_all_se)
putexcel G8=matrix(cp_all_ll)
putexcel H8=matrix(cp_all_ul)
putexcel I8=matrix(mcp_all_percent)
putexcel J8=matrix(mcp_all_se)
putexcel K8=matrix(mcp_all_ll)
putexcel L8=matrix(mcp_all_ul)
putexcel M8=matrix(unmettot_all_percent)
putexcel N8=matrix(unmettot_all_se)
putexcel O8=matrix(unmettot_all_ll)
putexcel P8=matrix(unmettot_all_ul)
putexcel R8=matrix(cp_mar_percent)
putexcel S8=matrix(cp_mar_se)
putexcel T8=matrix(cp_mar_ll)
putexcel U8=matrix(cp_mar_ul)
putexcel V8=matrix(mcp_mar_percent)
putexcel W8=matrix(mcp_mar_se)
putexcel X8=matrix(mcp_mar_ll)
putexcel Y8=matrix(mcp_mar_ul)
putexcel Z8=matrix(unmettot_mar_percent)
putexcel AA8=matrix(unmettot_mar_se)
putexcel AB8=matrix(unmettot_mar_ll)
putexcel AC8=matrix(unmettot_mar_ul)


***** ROUND 2
clear
use "~/Dropbox (Gates Institute)/01_Burkina/PMABF_Datasets/Round2/Final_PublicRelease/HHQ/PMA2015_BFR2_HHQFQ_v3_10Aug2018/PMA2015_BFR2_HHQFQ_v3_10Aug2018.dta", clear
putexcel B9=("Round 2")
putexcel C9=("4-6/2015")

** COUNT - Female Sample - All / Married Women  **
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel D9=matrix(FQresponse_1)
restore
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3) & (FQmarital_status==1 | FQmarital_status==2)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel Q9=matrix(FQresponse_1)
restore 

** Estimate Percentage and 95% CI
keep if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
egen all=tag(FQmetainstanceID)
egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)

svyset EA_ID [pw=FQweight], strata(strata) singleunit(scaled)
foreach group in all mar {
preserve
	keep if `group'==1
	foreach indicator in cp mcp unmettot {
		svy: prop `indicator', citype(wilson)
		matrix reference=r(table)
		matrix `indicator'_`group'_percent=reference[1,2]*100	
		matrix `indicator'_`group'_se=reference[2,2]*100
		matrix `indicator'_`group'_ll=reference[5,2]*100
		matrix `indicator'_`group'_ul=reference[6,2]*100
		}	
	restore
	}
putexcel E9=matrix(cp_all_percent)
putexcel F9=matrix(cp_all_se)
putexcel G9=matrix(cp_all_ll)
putexcel H9=matrix(cp_all_ul)
putexcel I9=matrix(mcp_all_percent)
putexcel J9=matrix(mcp_all_se)
putexcel K9=matrix(mcp_all_ll)
putexcel L9=matrix(mcp_all_ul)
putexcel M9=matrix(unmettot_all_percent)
putexcel N9=matrix(unmettot_all_se)
putexcel O9=matrix(unmettot_all_ll)
putexcel P9=matrix(unmettot_all_ul)
putexcel R9=matrix(cp_mar_percent)
putexcel S9=matrix(cp_mar_se)
putexcel T9=matrix(cp_mar_ll)
putexcel U9=matrix(cp_mar_ul)
putexcel V9=matrix(mcp_mar_percent)
putexcel W9=matrix(mcp_mar_se)
putexcel X9=matrix(mcp_mar_ll)
putexcel Y9=matrix(mcp_mar_ul)
putexcel Z9=matrix(unmettot_mar_percent)
putexcel AA9=matrix(unmettot_mar_se)
putexcel AB9=matrix(unmettot_mar_ll)
putexcel AC9=matrix(unmettot_mar_ul)


***** ROUND 3
clear
use "~/Dropbox (Gates Institute)/01_Burkina/PMABF_Datasets/Round3/Final_PublicRelease/HHQ/PMA2016_BFR3_HHQFQ_v3_10Aug2018/PMA2016_BFR3_HHQFQ_v3_10Aug2018.dta", clear
putexcel B10=("Round 3")
putexcel C10=("3-5/2015")
	
** COUNT - Female Sample - All / Married Women  **
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel D10=matrix(FQresponse_1)
restore
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3) & (FQmarital_status==1 | FQmarital_status==2)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel Q10=matrix(FQresponse_1)
restore 

** Estimate Percentage and 95% CI
keep if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
egen all=tag(FQmetainstanceID)
egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)

svyset EA_ID [pw=FQweight], strata(strata) singleunit(scaled)
foreach group in all mar {
preserve
	keep if `group'==1
	foreach indicator in cp mcp unmettot {
		svy: prop `indicator', citype(wilson)
		matrix reference=r(table)
		matrix `indicator'_`group'_percent=reference[1,2]*100	
		matrix `indicator'_`group'_se=reference[2,2]*100
		matrix `indicator'_`group'_ll=reference[5,2]*100
		matrix `indicator'_`group'_ul=reference[6,2]*100
		}	
	restore
	}
putexcel E10=matrix(cp_all_percent)
putexcel F10=matrix(cp_all_se)
putexcel G10=matrix(cp_all_ll)
putexcel H10=matrix(cp_all_ul)
putexcel I10=matrix(mcp_all_percent)
putexcel J10=matrix(mcp_all_se)
putexcel K10=matrix(mcp_all_ll)
putexcel L10=matrix(mcp_all_ul)
putexcel M10=matrix(unmettot_all_percent)
putexcel N10=matrix(unmettot_all_se)
putexcel O10=matrix(unmettot_all_ll)
putexcel P10=matrix(unmettot_all_ul)
putexcel R10=matrix(cp_mar_percent)
putexcel S10=matrix(cp_mar_se)
putexcel T10=matrix(cp_mar_ll)
putexcel U10=matrix(cp_mar_ul)
putexcel V10=matrix(mcp_mar_percent)
putexcel W10=matrix(mcp_mar_se)
putexcel X10=matrix(mcp_mar_ll)
putexcel Y10=matrix(mcp_mar_ul)
putexcel Z10=matrix(unmettot_mar_percent)
putexcel AA10=matrix(unmettot_mar_se)
putexcel AB10=matrix(unmettot_mar_ll)
putexcel AC10=matrix(unmettot_mar_ul)


**** ROUND 4
clear
use "~/Dropbox (Gates Institute)/01_Burkina/PMABF_Datasets/Round4/Final_PublicRelease/HHQ/PMA2016_BFR4_HHQFQ_v4_10Aug2018/PMA2016_BFR4_HHQFQ_v4_10Aug2018.dta", clear
putexcel B11=("Round 4")
putexcel C11=("12/2016-01/2017")

** COUNT - Female Sample - All / Married Women  **
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & last_night==1
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel D11=matrix(FQresponse_1)
restore
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & last_night==1 & (FQmarital_status==1 | FQmarital_status==2)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel Q11=matrix(FQresponse_1)
restore 

** Estimate Percentage and 95% CI
keep if FRS_result==1 & HHQ_result==1 & last_night==1
egen all=tag(FQmetainstanceID)
egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)

svyset EA_ID [pw=FQweight], strata(strata) singleunit(scaled)
foreach group in all mar {
preserve
	keep if `group'==1
	foreach indicator in cp mcp unmettot {
		svy: prop `indicator', citype(wilson)
		matrix reference=r(table)
		matrix `indicator'_`group'_percent=reference[1,2]*100	
		matrix `indicator'_`group'_se=reference[2,2]*100
		matrix `indicator'_`group'_ll=reference[5,2]*100
		matrix `indicator'_`group'_ul=reference[6,2]*100
		}	
	restore
	}
putexcel E11=matrix(cp_all_percent)
putexcel F11=matrix(cp_all_se)
putexcel G11=matrix(cp_all_ll)
putexcel H11=matrix(cp_all_ul)
putexcel I11=matrix(mcp_all_percent)
putexcel J11=matrix(mcp_all_se)
putexcel K11=matrix(mcp_all_ll)
putexcel L11=matrix(mcp_all_ul)
putexcel M11=matrix(unmettot_all_percent)
putexcel N11=matrix(unmettot_all_se)
putexcel O11=matrix(unmettot_all_ll)
putexcel P11=matrix(unmettot_all_ul)
putexcel R11=matrix(cp_mar_percent)
putexcel S11=matrix(cp_mar_se)
putexcel T11=matrix(cp_mar_ll)
putexcel U11=matrix(cp_mar_ul)
putexcel V11=matrix(mcp_mar_percent)
putexcel W11=matrix(mcp_mar_se)
putexcel X11=matrix(mcp_mar_ll)
putexcel Y11=matrix(mcp_mar_ul)
putexcel Z11=matrix(unmettot_mar_percent)
putexcel AA11=matrix(unmettot_mar_se)
putexcel AB11=matrix(unmettot_mar_ll)
putexcel AC11=matrix(unmettot_mar_ul)


**** ROUND 5
clear
*use "~/Dropbox (Gates Institute)/01_Burkina/PMABF_Datasets/Round5/Final_PublicRelease/HHQ/PMA2014_BFR1_HHQFQ_v3_10Aug2018/PMA2014_BFR1_HHQFQ_v3_10Aug2018.dta", clear
use "~/Dropbox (Gates Institute)/01_Burkina/PMABF_Datasets/Round5/Prelim100/BFR5_WealthWeightAll_2Jul2018.dta", clear
putexcel B12=("Round 5")
putexcel C12=("11-12/2017")

** COUNT - Female Sample - All / Married Women  **
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & last_night==1
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel D12=matrix(FQresponse_1)
restore
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & last_night==1 & (FQmarital_status==1 | FQmarital_status==2)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel Q12=matrix(FQresponse_1)
restore 

** Estimate Percentage and 95% CI
keep if FRS_result==1 & HHQ_result==1 & last_night==1
egen all=tag(FQmetainstanceID)
egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)

svyset EA [pw=FQweight], strata(strata) singleunit(scaled)
foreach group in all mar {
preserve
	keep if `group'==1
	foreach indicator in cp mcp unmettot {
		svy: prop `indicator', citype(wilson)
		matrix reference=r(table)
		matrix `indicator'_`group'_percent=reference[1,2]*100	
		matrix `indicator'_`group'_se=reference[2,2]*100
		matrix `indicator'_`group'_ll=reference[5,2]*100
		matrix `indicator'_`group'_ul=reference[6,2]*100
		}	
	restore
	}
putexcel E12=matrix(cp_all_percent)
putexcel F12=matrix(cp_all_se)
putexcel G12=matrix(cp_all_ll)
putexcel H12=matrix(cp_all_ul)
putexcel I12=matrix(mcp_all_percent)
putexcel J12=matrix(mcp_all_se)
putexcel K12=matrix(mcp_all_ll)
putexcel L12=matrix(mcp_all_ul)
putexcel M12=matrix(unmettot_all_percent)
putexcel N12=matrix(unmettot_all_se)
putexcel O12=matrix(unmettot_all_ll)
putexcel P12=matrix(unmettot_all_ul)
putexcel R12=matrix(cp_mar_percent)
putexcel S12=matrix(cp_mar_se)
putexcel T12=matrix(cp_mar_ll)
putexcel U12=matrix(cp_mar_ul)
putexcel V12=matrix(mcp_mar_percent)
putexcel W12=matrix(mcp_mar_se)
putexcel X12=matrix(mcp_mar_ll)
putexcel Y12=matrix(mcp_mar_ul)
putexcel Z12=matrix(unmettot_mar_percent)
putexcel AA12=matrix(unmettot_mar_se)
putexcel AB12=matrix(unmettot_mar_ll)
putexcel AC12=matrix(unmettot_mar_ul)




***********************************************************************************
*** 		COTE D'IVOIRE
***********************************************************************************
capture clear all
local excel "CI_KeyIndicators.xlsx"
local excel_sheet "CPR, mCPR, unmet need"
putexcel set "`excel'", modify sheet("`excel_sheet'")

**** ROUND 1
clear
use "~/Dropbox (Gates Institute)/02_Cote d'Ivoire/PMACI_Datasets/Round1/Final_PublicRelease/HHQ/PMA2017_CIR1_HHQFQ_v1_11May2018/PMA2017_CIR1_HHQFQ_v1_11May2018.dta",clear
putexcel A8=("PMA2020")
putexcel B8=("Round 1")
putexcel C8=("08-10/2017")

** COUNT - Female Sample - All / Married Women  **
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & last_night==1
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel D8=matrix(FQresponse_1)
restore
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & last_night==1 & (FQmarital_status==1 | FQmarital_status==2)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel Q8=matrix(FQresponse_1)
restore 

*** Estimate Percentage and 95% CI
keep if FRS_result==1 & HHQ_result==1 & last_night==1
egen all=tag(FQmetainstanceID)
egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)

svyset EA_ID [pw=FQweight], strata(strata) singleunit(scaled)
foreach group in all mar {
preserve
	keep if `group'==1
	foreach indicator in cp mcp unmettot {
		svy: prop `indicator', citype(wilson)
		matrix reference=r(table)
		matrix `indicator'_`group'_percent=reference[1,2]*100	
		matrix `indicator'_`group'_se=reference[2,2]*100
		matrix `indicator'_`group'_ll=reference[5,2]*100
		matrix `indicator'_`group'_ul=reference[6,2]*100
		}	
	restore
	}
putexcel E8=matrix(cp_all_percent)
putexcel F8=matrix(cp_all_se)
putexcel G8=matrix(cp_all_ll)
putexcel H8=matrix(cp_all_ul)
putexcel I8=matrix(mcp_all_percent)
putexcel J8=matrix(mcp_all_se)
putexcel K8=matrix(mcp_all_ll)
putexcel L8=matrix(mcp_all_ul)
putexcel M8=matrix(unmettot_all_percent)
putexcel N8=matrix(unmettot_all_se)
putexcel O8=matrix(unmettot_all_ll)
putexcel P8=matrix(unmettot_all_ul)
putexcel R8=matrix(cp_mar_percent)
putexcel S8=matrix(cp_mar_se)
putexcel T8=matrix(cp_mar_ll)
putexcel U8=matrix(cp_mar_ul)
putexcel V8=matrix(mcp_mar_percent)
putexcel W8=matrix(mcp_mar_se)
putexcel X8=matrix(mcp_mar_ll)
putexcel Y8=matrix(mcp_mar_ul)
putexcel Z8=matrix(unmettot_mar_percent)
putexcel AA8=matrix(unmettot_mar_se)
putexcel AB8=matrix(unmettot_mar_ll)
putexcel AC8=matrix(unmettot_mar_ul)




***********************************************************************************
*** 		DRC - CONGO
***********************************************************************************
local excel "DRC_KeyIndicators.xlsx"
local excel_sheet "CPR, mCPR, unmet need"
putexcel set "`excel'", modify sheet("`excel_sheet'")


**********************************
*** 		DRC - KINSHASA
**********************************
***** ROUND 1
clear
use "~/Dropbox (Gates Institute)/03_DRC/PMADRC_Datasets/Round1/Final_PublicRelease/HHQ/PMA2013_CDR1_Kinshasa_HHQFQ_v2_17May2017/PMA2013_CDR1_Kinshasa_HHQFQ_v2_17May2017.dta", clear
putexcel A8=("PMA2020")
putexcel B8=("Round 1")
putexcel C8=("10/2013-01/2014")

** COUNT - Female Sample - All / Married Women  **
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel D8=matrix(FQresponse_1)
restore
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3) & (FQmarital_status==1 | FQmarital_status==2)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel Q8=matrix(FQresponse_1)
restore 

*** Estimate Percentage and 95% CI
keep if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
egen all=tag(FQmetainstanceID)
egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)

svyset EA_ID [pw=FQweight], singleunit(scaled)
foreach group in all mar {
preserve
	keep if `group'==1
	foreach indicator in cp mcp unmettot {
		svy: prop `indicator', citype(wilson)
		matrix reference=r(table)
		matrix `indicator'_`group'_percent=reference[1,2]*100	
		matrix `indicator'_`group'_se=reference[2,2]*100
		matrix `indicator'_`group'_ll=reference[5,2]*100
		matrix `indicator'_`group'_ul=reference[6,2]*100
		}	
	restore
	}
putexcel E8=matrix(cp_all_percent)
putexcel F8=matrix(cp_all_se)
putexcel G8=matrix(cp_all_ll)
putexcel H8=matrix(cp_all_ul)
putexcel I8=matrix(mcp_all_percent)
putexcel J8=matrix(mcp_all_se)
putexcel K8=matrix(mcp_all_ll)
putexcel L8=matrix(mcp_all_ul)
putexcel M8=matrix(unmettot_all_percent)
putexcel N8=matrix(unmettot_all_se)
putexcel O8=matrix(unmettot_all_ll)
putexcel P8=matrix(unmettot_all_ul)
putexcel R8=matrix(cp_mar_percent)
putexcel S8=matrix(cp_mar_se)
putexcel T8=matrix(cp_mar_ll)
putexcel U8=matrix(cp_mar_ul)
putexcel V8=matrix(mcp_mar_percent)
putexcel W8=matrix(mcp_mar_se)
putexcel X8=matrix(mcp_mar_ll)
putexcel Y8=matrix(mcp_mar_ul)
putexcel Z8=matrix(unmettot_mar_percent)
putexcel AA8=matrix(unmettot_mar_se)
putexcel AB8=matrix(unmettot_mar_ll)
putexcel AC8=matrix(unmettot_mar_ul)


***** ROUND 2
clear
use "~/Dropbox (Gates Institute)/03_DRC/PMADRC_Datasets/Round2/Final_PublicRelease/HHQ/PMA2014_CDR2_Kinshasa_HHQFQ_v1_31Dec2016/PMA2014_CDR2_Kinshasa_HHQFQ_v1_2Jan2017.dta", clear
putexcel B9=("Round 2")
putexcel C9=("8-9/2014")

** COUNT - Female Sample - All / Married Women  **
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel D9=matrix(FQresponse_1)
restore
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3) & (FQmarital_status==1 | FQmarital_status==2)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel Q9=matrix(FQresponse_1)
restore 

** Estimate Percentage and 95% CI
keep if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
egen all=tag(FQmetainstanceID)
egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)

svyset EA_ID [pw=FQweight], singleunit(scaled)
foreach group in all mar {
preserve
	keep if `group'==1
	foreach indicator in cp mcp unmettot {
		svy: prop `indicator', citype(wilson)
		matrix reference=r(table)
		matrix `indicator'_`group'_percent=reference[1,2]*100	
		matrix `indicator'_`group'_se=reference[2,2]*100
		matrix `indicator'_`group'_ll=reference[5,2]*100
		matrix `indicator'_`group'_ul=reference[6,2]*100
		}	
	restore
	}
putexcel E9=matrix(cp_all_percent)
putexcel F9=matrix(cp_all_se)
putexcel G9=matrix(cp_all_ll)
putexcel H9=matrix(cp_all_ul)
putexcel I9=matrix(mcp_all_percent)
putexcel J9=matrix(mcp_all_se)
putexcel K9=matrix(mcp_all_ll)
putexcel L9=matrix(mcp_all_ul)
putexcel M9=matrix(unmettot_all_percent)
putexcel N9=matrix(unmettot_all_se)
putexcel O9=matrix(unmettot_all_ll)
putexcel P9=matrix(unmettot_all_ul)
putexcel R9=matrix(cp_mar_percent)
putexcel S9=matrix(cp_mar_se)
putexcel T9=matrix(cp_mar_ll)
putexcel U9=matrix(cp_mar_ul)
putexcel V9=matrix(mcp_mar_percent)
putexcel W9=matrix(mcp_mar_se)
putexcel X9=matrix(mcp_mar_ll)
putexcel Y9=matrix(mcp_mar_ul)
putexcel Z9=matrix(unmettot_mar_percent)
putexcel AA9=matrix(unmettot_mar_se)
putexcel AB9=matrix(unmettot_mar_ll)
putexcel AC9=matrix(unmettot_mar_ul)


***** ROUND 3
clear
use "~/Dropbox (Gates Institute)/03_DRC/PMADRC_Datasets/Round3/Final_PublicRelease/HHQ/PMA2015_CDR3_Kinshasa_HHQFQ_v2_18Nov2017/PMA2015_CDR3_Kinshasa_HHQFQ_v2_18Nov2017.dta", clear
putexcel B10=("Round 3")
putexcel C10=("5-6/2015")
	
** COUNT - Female Sample - All / Married Women  **
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel D10=matrix(FQresponse_1)
restore
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3) & (FQmarital_status==1 | FQmarital_status==2)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel Q10=matrix(FQresponse_1)
restore 

** Estimate Percentage and 95% CI
keep if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
egen all=tag(FQmetainstanceID)
egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)

svyset EA_ID [pw=FQweight], singleunit(scaled)
foreach group in all mar {
preserve
	keep if `group'==1
	foreach indicator in cp mcp unmettot {
		svy: prop `indicator', citype(wilson)
		matrix reference=r(table)
		matrix `indicator'_`group'_percent=reference[1,2]*100	
		matrix `indicator'_`group'_se=reference[2,2]*100
		matrix `indicator'_`group'_ll=reference[5,2]*100
		matrix `indicator'_`group'_ul=reference[6,2]*100
		}	
	restore
	}
putexcel E10=matrix(cp_all_percent)
putexcel F10=matrix(cp_all_se)
putexcel G10=matrix(cp_all_ll)
putexcel H10=matrix(cp_all_ul)
putexcel I10=matrix(mcp_all_percent)
putexcel J10=matrix(mcp_all_se)
putexcel K10=matrix(mcp_all_ll)
putexcel L10=matrix(mcp_all_ul)
putexcel M10=matrix(unmettot_all_percent)
putexcel N10=matrix(unmettot_all_se)
putexcel O10=matrix(unmettot_all_ll)
putexcel P10=matrix(unmettot_all_ul)
putexcel R10=matrix(cp_mar_percent)
putexcel S10=matrix(cp_mar_se)
putexcel T10=matrix(cp_mar_ll)
putexcel U10=matrix(cp_mar_ul)
putexcel V10=matrix(mcp_mar_percent)
putexcel W10=matrix(mcp_mar_se)
putexcel X10=matrix(mcp_mar_ll)
putexcel Y10=matrix(mcp_mar_ul)
putexcel Z10=matrix(unmettot_mar_percent)
putexcel AA10=matrix(unmettot_mar_se)
putexcel AB10=matrix(unmettot_mar_ll)
putexcel AC10=matrix(unmettot_mar_ul)


**** ROUND 4
clear
use "~/Dropbox (Gates Institute)/03_DRC/PMADRC_Datasets/Round4/Final_PublicRelease/HHQ/PMA2015_CDR4_Kinshasa_KongoCentral_HHQFQ_v2_28Jun2017/PMA2015_CDR4_Kinshasa_HHQFQ_v2_28Jun2017.dta", clear
putexcel B11=("Round 4")
putexcel C11=("11/2015-01/2016")

** COUNT - Female Sample - All / Married Women  **
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel D11=matrix(FQresponse_1)
restore
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3) & (FQmarital_status==1 | FQmarital_status==2)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel Q11=matrix(FQresponse_1)
restore 

** Estimate Percentage and 95% CI
keep if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
egen all=tag(FQmetainstanceID)
egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)

svyset EA_ID [pw=FQweight], singleunit(scaled)
foreach group in all mar {
preserve
	keep if `group'==1
	foreach indicator in cp mcp unmettot {
		svy: prop `indicator', citype(wilson)
		matrix reference=r(table)
		matrix `indicator'_`group'_percent=reference[1,2]*100	
		matrix `indicator'_`group'_se=reference[2,2]*100
		matrix `indicator'_`group'_ll=reference[5,2]*100
		matrix `indicator'_`group'_ul=reference[6,2]*100
		}	
	restore
	}
putexcel E11=matrix(cp_all_percent)
putexcel F11=matrix(cp_all_se)
putexcel G11=matrix(cp_all_ll)
putexcel H11=matrix(cp_all_ul)
putexcel I11=matrix(mcp_all_percent)
putexcel J11=matrix(mcp_all_se)
putexcel K11=matrix(mcp_all_ll)
putexcel L11=matrix(mcp_all_ul)
putexcel M11=matrix(unmettot_all_percent)
putexcel N11=matrix(unmettot_all_se)
putexcel O11=matrix(unmettot_all_ll)
putexcel P11=matrix(unmettot_all_ul)
putexcel R11=matrix(cp_mar_percent)
putexcel S11=matrix(cp_mar_se)
putexcel T11=matrix(cp_mar_ll)
putexcel U11=matrix(cp_mar_ul)
putexcel V11=matrix(mcp_mar_percent)
putexcel W11=matrix(mcp_mar_se)
putexcel X11=matrix(mcp_mar_ll)
putexcel Y11=matrix(mcp_mar_ul)
putexcel Z11=matrix(unmettot_mar_percent)
putexcel AA11=matrix(unmettot_mar_se)
putexcel AB11=matrix(unmettot_mar_ll)
putexcel AC11=matrix(unmettot_mar_ul)


**** ROUND 5
clear
use "~/Dropbox (Gates Institute)/03_DRC/PMADRC_Datasets/Round5/Final_PublicRelease/HHQ/PMA2016_CDR5_Kinshasa_KongoCentral_HHQFQ_v2_17Jan2018/PMA2016_CDR5_Kinshasa_HHQFQ_v2_17Jan2018.dta", clear
putexcel B12=("Round 5")
putexcel C12=("8-9/2016")

** COUNT - Female Sample - All / Married Women  **
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel D12=matrix(FQresponse_1)
restore
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3) & (FQmarital_status==1 | FQmarital_status==2)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel Q12=matrix(FQresponse_1)
restore 

** Estimate Percentage and 95% CI
keep if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
egen all=tag(FQmetainstanceID)
egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)

svyset EA_ID [pw=FQweight], singleunit(scaled)
foreach group in all mar {
preserve
	keep if `group'==1
	foreach indicator in cp mcp unmettot {
		svy: prop `indicator', citype(wilson)
		matrix reference=r(table)
		matrix `indicator'_`group'_percent=reference[1,2]*100	
		matrix `indicator'_`group'_se=reference[2,2]*100
		matrix `indicator'_`group'_ll=reference[5,2]*100
		matrix `indicator'_`group'_ul=reference[6,2]*100
		}	
	restore
	}
putexcel E12=matrix(cp_all_percent)
putexcel F12=matrix(cp_all_se)
putexcel G12=matrix(cp_all_ll)
putexcel H12=matrix(cp_all_ul)
putexcel I12=matrix(mcp_all_percent)
putexcel J12=matrix(mcp_all_se)
putexcel K12=matrix(mcp_all_ll)
putexcel L12=matrix(mcp_all_ul)
putexcel M12=matrix(unmettot_all_percent)
putexcel N12=matrix(unmettot_all_se)
putexcel O12=matrix(unmettot_all_ll)
putexcel P12=matrix(unmettot_all_ul)
putexcel R12=matrix(cp_mar_percent)
putexcel S12=matrix(cp_mar_se)
putexcel T12=matrix(cp_mar_ll)
putexcel U12=matrix(cp_mar_ul)
putexcel V12=matrix(mcp_mar_percent)
putexcel W12=matrix(mcp_mar_se)
putexcel X12=matrix(mcp_mar_ll)
putexcel Y12=matrix(mcp_mar_ul)
putexcel Z12=matrix(unmettot_mar_percent)
putexcel AA12=matrix(unmettot_mar_se)
putexcel AB12=matrix(unmettot_mar_ll)
putexcel AC12=matrix(unmettot_mar_ul)


**** ROUND 6
clear
use "~/Dropbox (Gates Institute)/03_DRC/PMADRC_Datasets/Round6/Final_PublicRelease/HHQFQ/PMA2017_CDR6_Kinshasa_KongoCentral_HHQFQ_v1_9Jul2018/PMA2017_CDR6_Kinshasa_HHQFQ_v1_9Jul2018.dta", clear
putexcel B13=("Round 6")
putexcel C13=("9-11/2017")

** COUNT - Female Sample - All / Married Women  **
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & last_night==1
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel D13=matrix(FQresponse_1)
restore
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & last_night==1 & (FQmarital_status==1 | FQmarital_status==2)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel Q13=matrix(FQresponse_1)
restore 

** Estimate Percentage and 95% CI
keep if FRS_result==1 & HHQ_result==1 & last_night==1
egen all=tag(FQmetainstanceID)
egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)

svyset EA_ID [pw=FQweight], singleunit(scaled)
foreach group in all mar {
preserve
	keep if `group'==1
	foreach indicator in cp mcp unmettot {
		svy: prop `indicator', citype(wilson)
		matrix reference=r(table)
		matrix `indicator'_`group'_percent=reference[1,2]*100	
		matrix `indicator'_`group'_se=reference[2,2]*100
		matrix `indicator'_`group'_ll=reference[5,2]*100
		matrix `indicator'_`group'_ul=reference[6,2]*100
		}	
	restore
	}
putexcel E13=matrix(cp_all_percent)
putexcel F13=matrix(cp_all_se)
putexcel G13=matrix(cp_all_ll)
putexcel H13=matrix(cp_all_ul)
putexcel I13=matrix(mcp_all_percent)
putexcel J13=matrix(mcp_all_se)
putexcel K13=matrix(mcp_all_ll)
putexcel L13=matrix(mcp_all_ul)
putexcel M13=matrix(unmettot_all_percent)
putexcel N13=matrix(unmettot_all_se)
putexcel O13=matrix(unmettot_all_ll)
putexcel P13=matrix(unmettot_all_ul)
putexcel R13=matrix(cp_mar_percent)
putexcel S13=matrix(cp_mar_se)
putexcel T13=matrix(cp_mar_ll)
putexcel U13=matrix(cp_mar_ul)
putexcel V13=matrix(mcp_mar_percent)
putexcel W13=matrix(mcp_mar_se)
putexcel X13=matrix(mcp_mar_ll)
putexcel Y13=matrix(mcp_mar_ul)
putexcel Z13=matrix(unmettot_mar_percent)
putexcel AA13=matrix(unmettot_mar_se)
putexcel AB13=matrix(unmettot_mar_ll)
putexcel AC13=matrix(unmettot_mar_ul)


**********************************
*** 		DRC - KONGO CENTRAL
**********************************
**** ROUND 4
clear
use "~/Dropbox (Gates Institute)/03_DRC/PMADRC_Datasets/Round4/Final_PublicRelease/HHQ/PMA2015_CDR4_Kinshasa_KongoCentral_HHQFQ_v2_28Jun2017/PMA2015_CDR4_KongoCentral_HHQFQ_v2_28Jun2017.dta", clear
putexcel B15=("Round 4")
putexcel C15=("11/2015-01/2016")

** COUNT - Female Sample - All / Married Women  **
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel D15=matrix(FQresponse_1)
restore
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3) & (FQmarital_status==1 | FQmarital_status==2)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel Q15=matrix(FQresponse_1)
restore 

** Estimate Percentage and 95% CI
keep if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
egen all=tag(FQmetainstanceID)
egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)

svyset EA_ID [pw=FQweight], singleunit(scaled)
foreach group in all mar {
preserve
	keep if `group'==1
	foreach indicator in cp mcp unmettot {
		svy: prop `indicator', citype(wilson)
		matrix reference=r(table)
		matrix `indicator'_`group'_percent=reference[1,2]*100	
		matrix `indicator'_`group'_se=reference[2,2]*100
		matrix `indicator'_`group'_ll=reference[5,2]*100
		matrix `indicator'_`group'_ul=reference[6,2]*100
		}	
	restore
	}
putexcel E15=matrix(cp_all_percent)
putexcel F15=matrix(cp_all_se)
putexcel G15=matrix(cp_all_ll)
putexcel H15=matrix(cp_all_ul)
putexcel I15=matrix(mcp_all_percent)
putexcel J15=matrix(mcp_all_se)
putexcel K15=matrix(mcp_all_ll)
putexcel L15=matrix(mcp_all_ul)
putexcel M15=matrix(unmettot_all_percent)
putexcel N15=matrix(unmettot_all_se)
putexcel O15=matrix(unmettot_all_ll)
putexcel P15=matrix(unmettot_all_ul)
putexcel R15=matrix(cp_mar_percent)
putexcel S15=matrix(cp_mar_se)
putexcel T15=matrix(cp_mar_ll)
putexcel U15=matrix(cp_mar_ul)
putexcel V15=matrix(mcp_mar_percent)
putexcel W15=matrix(mcp_mar_se)
putexcel X15=matrix(mcp_mar_ll)
putexcel Y15=matrix(mcp_mar_ul)
putexcel Z15=matrix(unmettot_mar_percent)
putexcel AA15=matrix(unmettot_mar_se)
putexcel AB15=matrix(unmettot_mar_ll)
putexcel AC15=matrix(unmettot_mar_ul)


**** ROUND 5
clear
use "~/Dropbox (Gates Institute)/03_DRC/PMADRC_Datasets/Round5/Final_PublicRelease/HHQ/PMA2016_CDR5_Kinshasa_KongoCentral_HHQFQ_v2_17Jan2018/PMA2016_CDR5_KongoCentral_HHQFQ_v2_17Jan2018.dta", clear
putexcel B16=("Round 5")
putexcel C16=("8-9/2016")

** COUNT - Female Sample - All / Married Women  **
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel D16=matrix(FQresponse_1)
restore
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3) & (FQmarital_status==1 | FQmarital_status==2)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel Q16=matrix(FQresponse_1)
restore 

** Estimate Percentage and 95% CI
keep if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
egen all=tag(FQmetainstanceID)
egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)

svyset EA_ID [pw=FQweight], singleunit(scaled)
foreach group in all mar {
preserve
	keep if `group'==1
	foreach indicator in cp mcp unmettot {
		svy: prop `indicator', citype(wilson)
		matrix reference=r(table)
		matrix `indicator'_`group'_percent=reference[1,2]*100	
		matrix `indicator'_`group'_se=reference[2,2]*100
		matrix `indicator'_`group'_ll=reference[5,2]*100
		matrix `indicator'_`group'_ul=reference[6,2]*100
		}	
	restore
	}
putexcel E16=matrix(cp_all_percent)
putexcel F16=matrix(cp_all_se)
putexcel G16=matrix(cp_all_ll)
putexcel H16=matrix(cp_all_ul)
putexcel I16=matrix(mcp_all_percent)
putexcel J16=matrix(mcp_all_se)
putexcel K16=matrix(mcp_all_ll)
putexcel L16=matrix(mcp_all_ul)
putexcel M16=matrix(unmettot_all_percent)
putexcel N16=matrix(unmettot_all_se)
putexcel O16=matrix(unmettot_all_ll)
putexcel P16=matrix(unmettot_all_ul)
putexcel R16=matrix(cp_mar_percent)
putexcel S16=matrix(cp_mar_se)
putexcel T16=matrix(cp_mar_ll)
putexcel U16=matrix(cp_mar_ul)
putexcel V16=matrix(mcp_mar_percent)
putexcel W16=matrix(mcp_mar_se)
putexcel X16=matrix(mcp_mar_ll)
putexcel Y16=matrix(mcp_mar_ul)
putexcel Z16=matrix(unmettot_mar_percent)
putexcel AA16=matrix(unmettot_mar_se)
putexcel AB16=matrix(unmettot_mar_ll)
putexcel AC16=matrix(unmettot_mar_ul)


**** ROUND 6
clear
use "~/Dropbox (Gates Institute)/03_DRC/PMADRC_Datasets/Round6/Final_PublicRelease/HHQFQ/PMA2017_CDR6_Kinshasa_KongoCentral_HHQFQ_v1_9Jul2018/PMA2017_CDR6_KongoCentral_HHQFQ_v1_9Jul2018.dta", clear
putexcel B17=("Round 6")
putexcel C17=("9-11/2017")

** COUNT - Female Sample - All / Married Women  **
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & last_night==1
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel D17=matrix(FQresponse_1)
restore
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & last_night==1 & (FQmarital_status==1 | FQmarital_status==2)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel Q17=matrix(FQresponse_1)
restore 

** Estimate Percentage and 95% CI
keep if FRS_result==1 & HHQ_result==1 & last_night==1
egen all=tag(FQmetainstanceID)
egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)

svyset EA_ID [pw=FQweight], singleunit(scaled)
foreach group in all mar {
preserve
	keep if `group'==1
	foreach indicator in cp mcp unmettot {
		svy: prop `indicator', citype(wilson)
		matrix reference=r(table)
		matrix `indicator'_`group'_percent=reference[1,2]*100	
		matrix `indicator'_`group'_se=reference[2,2]*100
		matrix `indicator'_`group'_ll=reference[5,2]*100
		matrix `indicator'_`group'_ul=reference[6,2]*100
		}	
	restore
	}
putexcel E17=matrix(cp_all_percent)
putexcel F17=matrix(cp_all_se)
putexcel G17=matrix(cp_all_ll)
putexcel H17=matrix(cp_all_ul)
putexcel I17=matrix(mcp_all_percent)
putexcel J17=matrix(mcp_all_se)
putexcel K17=matrix(mcp_all_ll)
putexcel L17=matrix(mcp_all_ul)
putexcel M17=matrix(unmettot_all_percent)
putexcel N17=matrix(unmettot_all_se)
putexcel O17=matrix(unmettot_all_ll)
putexcel P17=matrix(unmettot_all_ul)
putexcel R17=matrix(cp_mar_percent)
putexcel S17=matrix(cp_mar_se)
putexcel T17=matrix(cp_mar_ll)
putexcel U17=matrix(cp_mar_ul)
putexcel V17=matrix(mcp_mar_percent)
putexcel W17=matrix(mcp_mar_se)
putexcel X17=matrix(mcp_mar_ll)
putexcel Y17=matrix(mcp_mar_ul)
putexcel Z17=matrix(unmettot_mar_percent)
putexcel AA17=matrix(unmettot_mar_se)
putexcel AB17=matrix(unmettot_mar_ll)
putexcel AC17=matrix(unmettot_mar_ul)




***********************************************************************************
*** 		ETHIOPIA
***********************************************************************************
capture clear all
local excel "ETH_KeyIndicators.xlsx"
local excel_sheet "CPR, mCPR, unmet need"
putexcel set "`excel'", modify sheet("`excel_sheet'")

***** ROUND 1
clear
use "~/Dropbox (Gates Institute)/04_Ethiopia/PMAET_Datasets/Round1/Final_PublicRelease/HHQ/PMA2014_ETR1_HHQFQ_v4_13Aug2018/PMA2014_ETR1_HHQFQ_v4_13Aug2018.dta", clear
putexcel A8=("PMA2020")
putexcel B8=("Round 1")
putexcel C8=("1-3/2014")

** COUNT - Female Sample - All / Married Women  **
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel D8=matrix(FQresponse_1)
restore
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3) & (FQmarital_status==1 | FQmarital_status==2)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel Q8=matrix(FQresponse_1)
restore 

*** Estimate Percentage and 95% CI
keep if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
egen all=tag(FQmetainstanceID)
egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)

svyset EA_ID [pw=FQweight], strata(strata) singleunit(scaled)
foreach group in all mar {
preserve
	keep if `group'==1
	foreach indicator in cp mcp unmettot {
		svy: prop `indicator', citype(wilson)
		matrix reference=r(table)
		matrix `indicator'_`group'_percent=reference[1,2]*100	
		matrix `indicator'_`group'_se=reference[2,2]*100
		matrix `indicator'_`group'_ll=reference[5,2]*100
		matrix `indicator'_`group'_ul=reference[6,2]*100
		}	
	restore
	}
putexcel E8=matrix(cp_all_percent)
putexcel F8=matrix(cp_all_se)
putexcel G8=matrix(cp_all_ll)
putexcel H8=matrix(cp_all_ul)
putexcel I8=matrix(mcp_all_percent)
putexcel J8=matrix(mcp_all_se)
putexcel K8=matrix(mcp_all_ll)
putexcel L8=matrix(mcp_all_ul)
putexcel M8=matrix(unmettot_all_percent)
putexcel N8=matrix(unmettot_all_se)
putexcel O8=matrix(unmettot_all_ll)
putexcel P8=matrix(unmettot_all_ul)
putexcel R8=matrix(cp_mar_percent)
putexcel S8=matrix(cp_mar_se)
putexcel T8=matrix(cp_mar_ll)
putexcel U8=matrix(cp_mar_ul)
putexcel V8=matrix(mcp_mar_percent)
putexcel W8=matrix(mcp_mar_se)
putexcel X8=matrix(mcp_mar_ll)
putexcel Y8=matrix(mcp_mar_ul)
putexcel Z8=matrix(unmettot_mar_percent)
putexcel AA8=matrix(unmettot_mar_se)
putexcel AB8=matrix(unmettot_mar_ll)
putexcel AC8=matrix(unmettot_mar_ul)


***** ROUND 2
clear
use "~/Dropbox (Gates Institute)/04_Ethiopia/PMAET_Datasets/Round2/Final_PublicRelease/HHQ/PMA2014_ETR2_HHQFQ_v2_13Aug2018/PMA2014_ETR2_HHQFQ_v2_13Aug2018.dta", clear
putexcel B9=("Round 2")
putexcel C9=("10-12/2014")

** COUNT - Female Sample - All / Married Women  **
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel D9=matrix(FQresponse_1)
restore
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3) & (FQmarital_status==1 | FQmarital_status==2)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel Q9=matrix(FQresponse_1)
restore 

** Estimate Percentage and 95% CI
keep if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
egen all=tag(FQmetainstanceID)
egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)

svyset EA_ID [pw=FQweight], strata(strata) singleunit(scaled)
foreach group in all mar {
preserve
	keep if `group'==1
	foreach indicator in cp mcp unmettot {
		svy: prop `indicator', citype(wilson)
		matrix reference=r(table)
		matrix `indicator'_`group'_percent=reference[1,2]*100	
		matrix `indicator'_`group'_se=reference[2,2]*100
		matrix `indicator'_`group'_ll=reference[5,2]*100
		matrix `indicator'_`group'_ul=reference[6,2]*100
		}	
	restore
	}
putexcel E9=matrix(cp_all_percent)
putexcel F9=matrix(cp_all_se)
putexcel G9=matrix(cp_all_ll)
putexcel H9=matrix(cp_all_ul)
putexcel I9=matrix(mcp_all_percent)
putexcel J9=matrix(mcp_all_se)
putexcel K9=matrix(mcp_all_ll)
putexcel L9=matrix(mcp_all_ul)
putexcel M9=matrix(unmettot_all_percent)
putexcel N9=matrix(unmettot_all_se)
putexcel O9=matrix(unmettot_all_ll)
putexcel P9=matrix(unmettot_all_ul)
putexcel R9=matrix(cp_mar_percent)
putexcel S9=matrix(cp_mar_se)
putexcel T9=matrix(cp_mar_ll)
putexcel U9=matrix(cp_mar_ul)
putexcel V9=matrix(mcp_mar_percent)
putexcel W9=matrix(mcp_mar_se)
putexcel X9=matrix(mcp_mar_ll)
putexcel Y9=matrix(mcp_mar_ul)
putexcel Z9=matrix(unmettot_mar_percent)
putexcel AA9=matrix(unmettot_mar_se)
putexcel AB9=matrix(unmettot_mar_ll)
putexcel AC9=matrix(unmettot_mar_ul)


***** ROUND 3
clear
use "~/Dropbox (Gates Institute)/04_Ethiopia/PMAET_Datasets/Round3/Final_PublicRelease/HHQ/PMA2015_ETR3_HHQFQ_v2_13Aug2018/PMA2015_ETR3_HHQFQ_v2_13Aug2018.dta", clear
putexcel B10=("Round 3")
putexcel C10=("4-5/2015")
	
** COUNT - Female Sample - All / Married Women  **
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel D10=matrix(FQresponse_1)
restore
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3) & (FQmarital_status==1 | FQmarital_status==2)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel Q10=matrix(FQresponse_1)
restore 

** Estimate Percentage and 95% CI
keep if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
egen all=tag(FQmetainstanceID)
egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)

svyset EA_ID [pw=FQweight], strata(strata) singleunit(scaled)
foreach group in all mar {
preserve
	keep if `group'==1
	foreach indicator in cp mcp unmettot {
		svy: prop `indicator', citype(wilson)
		matrix reference=r(table)
		matrix `indicator'_`group'_percent=reference[1,2]*100	
		matrix `indicator'_`group'_se=reference[2,2]*100
		matrix `indicator'_`group'_ll=reference[5,2]*100
		matrix `indicator'_`group'_ul=reference[6,2]*100
		}	
	restore
	}
putexcel E10=matrix(cp_all_percent)
putexcel F10=matrix(cp_all_se)
putexcel G10=matrix(cp_all_ll)
putexcel H10=matrix(cp_all_ul)
putexcel I10=matrix(mcp_all_percent)
putexcel J10=matrix(mcp_all_se)
putexcel K10=matrix(mcp_all_ll)
putexcel L10=matrix(mcp_all_ul)
putexcel M10=matrix(unmettot_all_percent)
putexcel N10=matrix(unmettot_all_se)
putexcel O10=matrix(unmettot_all_ll)
putexcel P10=matrix(unmettot_all_ul)
putexcel R10=matrix(cp_mar_percent)
putexcel S10=matrix(cp_mar_se)
putexcel T10=matrix(cp_mar_ll)
putexcel U10=matrix(cp_mar_ul)
putexcel V10=matrix(mcp_mar_percent)
putexcel W10=matrix(mcp_mar_se)
putexcel X10=matrix(mcp_mar_ll)
putexcel Y10=matrix(mcp_mar_ul)
putexcel Z10=matrix(unmettot_mar_percent)
putexcel AA10=matrix(unmettot_mar_se)
putexcel AB10=matrix(unmettot_mar_ll)
putexcel AC10=matrix(unmettot_mar_ul)


**** ROUND 4
clear
use "~/Dropbox (Gates Institute)/04_Ethiopia/PMAET_Datasets/Round4/Final_PublicRelease/HHQ/PMA2016_ETR4_HHQFQ_v2_13Aug2018/PMA2016_ETR4_HHQFQ_v2_13Aug2018.dta", clear
putexcel B11=("Round 4")
putexcel C11=("3-5/2016")

** COUNT - Female Sample - All / Married Women  **
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel D11=matrix(FQresponse_1)
restore
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3) & (FQmarital_status==1 | FQmarital_status==2)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel Q11=matrix(FQresponse_1)
restore 

** Estimate Percentage and 95% CI
keep if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
egen all=tag(FQmetainstanceID)
egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)

svyset EA_ID [pw=FQweight], strata(strata) singleunit(scaled)
foreach group in all mar {
preserve
	keep if `group'==1
	foreach indicator in cp mcp unmettot {
		svy: prop `indicator', citype(wilson)
		matrix reference=r(table)
		matrix `indicator'_`group'_percent=reference[1,2]*100	
		matrix `indicator'_`group'_se=reference[2,2]*100
		matrix `indicator'_`group'_ll=reference[5,2]*100
		matrix `indicator'_`group'_ul=reference[6,2]*100
		}	
	restore
	}
putexcel E11=matrix(cp_all_percent)
putexcel F11=matrix(cp_all_se)
putexcel G11=matrix(cp_all_ll)
putexcel H11=matrix(cp_all_ul)
putexcel I11=matrix(mcp_all_percent)
putexcel J11=matrix(mcp_all_se)
putexcel K11=matrix(mcp_all_ll)
putexcel L11=matrix(mcp_all_ul)
putexcel M11=matrix(unmettot_all_percent)
putexcel N11=matrix(unmettot_all_se)
putexcel O11=matrix(unmettot_all_ll)
putexcel P11=matrix(unmettot_all_ul)
putexcel R11=matrix(cp_mar_percent)
putexcel S11=matrix(cp_mar_se)
putexcel T11=matrix(cp_mar_ll)
putexcel U11=matrix(cp_mar_ul)
putexcel V11=matrix(mcp_mar_percent)
putexcel W11=matrix(mcp_mar_se)
putexcel X11=matrix(mcp_mar_ll)
putexcel Y11=matrix(mcp_mar_ul)
putexcel Z11=matrix(unmettot_mar_percent)
putexcel AA11=matrix(unmettot_mar_se)
putexcel AB11=matrix(unmettot_mar_ll)
putexcel AC11=matrix(unmettot_mar_ul)


**** ROUND 5
clear
use "~/Dropbox (Gates Institute)/04_Ethiopia/PMAET_Datasets/Round5/Final_PublicRelease/HHQ/PMA2017_ETR5_HHQFQ_v2_13Aug2018/PMA2017_ETR5_HHQFQ_v2_13Aug2018.dta", clear
putexcel B12=("Round 5")
putexcel C12=("4-5/2017")

** COUNT - Female Sample - All / Married Women  **
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & last_night==1
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel D12=matrix(FQresponse_1)
restore
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & last_night==1 & (FQmarital_status==1 | FQmarital_status==2)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel Q12=matrix(FQresponse_1)
restore 

** Estimate Percentage and 95% CI
keep if FRS_result==1 & HHQ_result==1 & last_night==1
egen all=tag(FQmetainstanceID)
egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)

svyset EA_ID [pw=FQweight], strata(strata) singleunit(scaled)
foreach group in all mar {
preserve
	keep if `group'==1
	foreach indicator in cp mcp unmettot {
		svy: prop `indicator', citype(wilson)
		matrix reference=r(table)
		matrix `indicator'_`group'_percent=reference[1,2]*100	
		matrix `indicator'_`group'_se=reference[2,2]*100
		matrix `indicator'_`group'_ll=reference[5,2]*100
		matrix `indicator'_`group'_ul=reference[6,2]*100
		}	
	restore
	}
putexcel E12=matrix(cp_all_percent)
putexcel F12=matrix(cp_all_se)
putexcel G12=matrix(cp_all_ll)
putexcel H12=matrix(cp_all_ul)
putexcel I12=matrix(mcp_all_percent)
putexcel J12=matrix(mcp_all_se)
putexcel K12=matrix(mcp_all_ll)
putexcel L12=matrix(mcp_all_ul)
putexcel M12=matrix(unmettot_all_percent)
putexcel N12=matrix(unmettot_all_se)
putexcel O12=matrix(unmettot_all_ll)
putexcel P12=matrix(unmettot_all_ul)
putexcel R12=matrix(cp_mar_percent)
putexcel S12=matrix(cp_mar_se)
putexcel T12=matrix(cp_mar_ll)
putexcel U12=matrix(cp_mar_ul)
putexcel V12=matrix(mcp_mar_percent)
putexcel W12=matrix(mcp_mar_se)
putexcel X12=matrix(mcp_mar_ll)
putexcel Y12=matrix(mcp_mar_ul)
putexcel Z12=matrix(unmettot_mar_percent)
putexcel AA12=matrix(unmettot_mar_se)
putexcel AB12=matrix(unmettot_mar_ll)
putexcel AC12=matrix(unmettot_mar_ul)




***********************************************************************************
*** 		GHANA
***********************************************************************************
capture clear all
local excel "GH_KeyIndicators.xlsx"
local excel_sheet "CPR, mCPR, unmet need"
putexcel set "`excel'", modify sheet("`excel_sheet'")

***** ROUND 1
clear
use "~/Dropbox (Gates Institute)/05_Ghana/PMAGH_Datasets/Round1/Final_PublicRelease/HHQFQ/PMA2013_GHR1_HHQFQ_v2_18Nov2017/PMA2013_GHR1_HHQFQ_v2_20Oct2017.dta", clear
putexcel A8=("PMA2020")
putexcel B8=("Round 1")
putexcel C8=("9-10/2013")

** COUNT - Female Sample - All / Married Women  **
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel D8=matrix(FQresponse_1)
restore
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3) & (FQmarital_status==1 | FQmarital_status==2)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel Q8=matrix(FQresponse_1)
restore 

*** Estimate Percentage and 95% CI
keep if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
egen all=tag(FQmetainstanceID)
egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)

svyset EA_ID [pw=FQweight], strata(strata) singleunit(scaled)
foreach group in all mar {
preserve
	keep if `group'==1
	foreach indicator in cp mcp unmettot {
		svy: prop `indicator', citype(wilson)
		matrix reference=r(table)
		matrix `indicator'_`group'_percent=reference[1,2]*100	
		matrix `indicator'_`group'_se=reference[2,2]*100
		matrix `indicator'_`group'_ll=reference[5,2]*100
		matrix `indicator'_`group'_ul=reference[6,2]*100
		}	
	restore
	}
putexcel E8=matrix(cp_all_percent)
putexcel F8=matrix(cp_all_se)
putexcel G8=matrix(cp_all_ll)
putexcel H8=matrix(cp_all_ul)
putexcel I8=matrix(mcp_all_percent)
putexcel J8=matrix(mcp_all_se)
putexcel K8=matrix(mcp_all_ll)
putexcel L8=matrix(mcp_all_ul)
putexcel M8=matrix(unmettot_all_percent)
putexcel N8=matrix(unmettot_all_se)
putexcel O8=matrix(unmettot_all_ll)
putexcel P8=matrix(unmettot_all_ul)
putexcel R8=matrix(cp_mar_percent)
putexcel S8=matrix(cp_mar_se)
putexcel T8=matrix(cp_mar_ll)
putexcel U8=matrix(cp_mar_ul)
putexcel V8=matrix(mcp_mar_percent)
putexcel W8=matrix(mcp_mar_se)
putexcel X8=matrix(mcp_mar_ll)
putexcel Y8=matrix(mcp_mar_ul)
putexcel Z8=matrix(unmettot_mar_percent)
putexcel AA8=matrix(unmettot_mar_se)
putexcel AB8=matrix(unmettot_mar_ll)
putexcel AC8=matrix(unmettot_mar_ul)


***** ROUND 2
clear
use "~/Dropbox (Gates Institute)/05_Ghana/PMAGH_Datasets/Round2/Final_PublicRelease/HHQFQ/PMA2014_GHR2_HHQFQ_v1_23Dec2016/PMA2014_GHR2_HHQFQ_v1_23Dec2016.dta", clear
putexcel B9=("Round 2")
putexcel C9=("1-2/2014")

** COUNT - Female Sample - All / Married Women  **
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel D9=matrix(FQresponse_1)
restore
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3) & (FQmarital_status==1 | FQmarital_status==2)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel Q9=matrix(FQresponse_1)
restore 

** Estimate Percentage and 95% CI
keep if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
egen all=tag(FQmetainstanceID)
egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)

svyset EA_ID [pw=FQweight], strata(strata) singleunit(scaled)
foreach group in all mar {
preserve
	keep if `group'==1
	foreach indicator in cp mcp unmettot {
		svy: prop `indicator', citype(wilson)
		matrix reference=r(table)
		matrix `indicator'_`group'_percent=reference[1,2]*100	
		matrix `indicator'_`group'_se=reference[2,2]*100
		matrix `indicator'_`group'_ll=reference[5,2]*100
		matrix `indicator'_`group'_ul=reference[6,2]*100
		}	
	restore
	}
putexcel E9=matrix(cp_all_percent)
putexcel F9=matrix(cp_all_se)
putexcel G9=matrix(cp_all_ll)
putexcel H9=matrix(cp_all_ul)
putexcel I9=matrix(mcp_all_percent)
putexcel J9=matrix(mcp_all_se)
putexcel K9=matrix(mcp_all_ll)
putexcel L9=matrix(mcp_all_ul)
putexcel M9=matrix(unmettot_all_percent)
putexcel N9=matrix(unmettot_all_se)
putexcel O9=matrix(unmettot_all_ll)
putexcel P9=matrix(unmettot_all_ul)
putexcel R9=matrix(cp_mar_percent)
putexcel S9=matrix(cp_mar_se)
putexcel T9=matrix(cp_mar_ll)
putexcel U9=matrix(cp_mar_ul)
putexcel V9=matrix(mcp_mar_percent)
putexcel W9=matrix(mcp_mar_se)
putexcel X9=matrix(mcp_mar_ll)
putexcel Y9=matrix(mcp_mar_ul)
putexcel Z9=matrix(unmettot_mar_percent)
putexcel AA9=matrix(unmettot_mar_se)
putexcel AB9=matrix(unmettot_mar_ll)
putexcel AC9=matrix(unmettot_mar_ul)


***** ROUND 3
clear
use "~/Dropbox (Gates Institute)/05_Ghana/PMAGH_Datasets/Round3/Final_PublicRelease/HHQ/PMA2014_UGR3_HHQFQ_v1_23Dec2016/PMA2014_UGR3_HHQFQ_v1_23Dec2016.dta", clear
putexcel B10=("Round 3")
putexcel C10=("9-12/2014")
	
** COUNT - Female Sample - All / Married Women  **
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel D10=matrix(FQresponse_1)
restore
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3) & (FQmarital_status==1 | FQmarital_status==2)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel Q10=matrix(FQresponse_1)
restore 

** Estimate Percentage and 95% CI
keep if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
egen all=tag(FQmetainstanceID)
egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)

svyset EA_ID [pw=FQweight], strata(strata) singleunit(scaled)
foreach group in all mar {
preserve
	keep if `group'==1
	foreach indicator in cp mcp unmettot {
		svy: prop `indicator', citype(wilson)
		matrix reference=r(table)
		matrix `indicator'_`group'_percent=reference[1,2]*100	
		matrix `indicator'_`group'_se=reference[2,2]*100
		matrix `indicator'_`group'_ll=reference[5,2]*100
		matrix `indicator'_`group'_ul=reference[6,2]*100
		}	
	restore
	}
putexcel E10=matrix(cp_all_percent)
putexcel F10=matrix(cp_all_se)
putexcel G10=matrix(cp_all_ll)
putexcel H10=matrix(cp_all_ul)
putexcel I10=matrix(mcp_all_percent)
putexcel J10=matrix(mcp_all_se)
putexcel K10=matrix(mcp_all_ll)
putexcel L10=matrix(mcp_all_ul)
putexcel M10=matrix(unmettot_all_percent)
putexcel N10=matrix(unmettot_all_se)
putexcel O10=matrix(unmettot_all_ll)
putexcel P10=matrix(unmettot_all_ul)
putexcel R10=matrix(cp_mar_percent)
putexcel S10=matrix(cp_mar_se)
putexcel T10=matrix(cp_mar_ll)
putexcel U10=matrix(cp_mar_ul)
putexcel V10=matrix(mcp_mar_percent)
putexcel W10=matrix(mcp_mar_se)
putexcel X10=matrix(mcp_mar_ll)
putexcel Y10=matrix(mcp_mar_ul)
putexcel Z10=matrix(unmettot_mar_percent)
putexcel AA10=matrix(unmettot_mar_se)
putexcel AB10=matrix(unmettot_mar_ll)
putexcel AC10=matrix(unmettot_mar_ul)


**** ROUND 4
clear
use "~/Dropbox (Gates Institute)/05_Ghana/PMAGH_Datasets/Round4/Final_PublicRelease/HHQ/PMA2015_GHR4_HHQFQ_v1_23Dec2016/PMA2015_GHR4_HHQFQ_v1_23Dec2016.dta", clear
putexcel B11=("Round 4")
putexcel C11=("5-6/2015")

** COUNT - Female Sample - All / Married Women  **
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel D11=matrix(FQresponse_1)
restore
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3) & (FQmarital_status==1 | FQmarital_status==2)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel Q11=matrix(FQresponse_1)
restore 

** Estimate Percentage and 95% CI
keep if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
egen all=tag(FQmetainstanceID)
egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)

svyset EA_ID [pw=FQweight], strata(strata) singleunit(scaled)
foreach group in all mar {
preserve
	keep if `group'==1
	foreach indicator in cp mcp unmettot {
		svy: prop `indicator', citype(wilson)
		matrix reference=r(table)
		matrix `indicator'_`group'_percent=reference[1,2]*100	
		matrix `indicator'_`group'_se=reference[2,2]*100
		matrix `indicator'_`group'_ll=reference[5,2]*100
		matrix `indicator'_`group'_ul=reference[6,2]*100
		}	
	restore
	}
putexcel E11=matrix(cp_all_percent)
putexcel F11=matrix(cp_all_se)
putexcel G11=matrix(cp_all_ll)
putexcel H11=matrix(cp_all_ul)
putexcel I11=matrix(mcp_all_percent)
putexcel J11=matrix(mcp_all_se)
putexcel K11=matrix(mcp_all_ll)
putexcel L11=matrix(mcp_all_ul)
putexcel M11=matrix(unmettot_all_percent)
putexcel N11=matrix(unmettot_all_se)
putexcel O11=matrix(unmettot_all_ll)
putexcel P11=matrix(unmettot_all_ul)
putexcel R11=matrix(cp_mar_percent)
putexcel S11=matrix(cp_mar_se)
putexcel T11=matrix(cp_mar_ll)
putexcel U11=matrix(cp_mar_ul)
putexcel V11=matrix(mcp_mar_percent)
putexcel W11=matrix(mcp_mar_se)
putexcel X11=matrix(mcp_mar_ll)
putexcel Y11=matrix(mcp_mar_ul)
putexcel Z11=matrix(unmettot_mar_percent)
putexcel AA11=matrix(unmettot_mar_se)
putexcel AB11=matrix(unmettot_mar_ll)
putexcel AC11=matrix(unmettot_mar_ul)


**** ROUND 5
clear
use "~/Dropbox (Gates Institute)/05_Ghana/PMAGH_Datasets/Round5/Final_PublicRelease/HHQ/PMA2016_GHR5_HHQFQ_v1_22Aug2017/PMA2016_GHR5_HHQFQ_v1_22Aug2017.dta", clear
putexcel B12=("Round 5")
putexcel C12=("8-11/2016")

** COUNT - Female Sample - All / Married Women  **
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & last_night==1
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel D12=matrix(FQresponse_1)
restore
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & last_night==1 & (FQmarital_status==1 | FQmarital_status==2)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel Q12=matrix(FQresponse_1)
restore 

** Estimate Percentage and 95% CI
keep if FRS_result==1 & HHQ_result==1 & last_night==1
egen all=tag(FQmetainstanceID)
egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)

svyset EA_ID [pw=FQweight], strata(strata) singleunit(scaled)
foreach group in all mar {
preserve
	keep if `group'==1
	foreach indicator in cp mcp unmettot {
		svy: prop `indicator', citype(wilson)
		matrix reference=r(table)
		matrix `indicator'_`group'_percent=reference[1,2]*100	
		matrix `indicator'_`group'_se=reference[2,2]*100
		matrix `indicator'_`group'_ll=reference[5,2]*100
		matrix `indicator'_`group'_ul=reference[6,2]*100
		}	
	restore
	}
putexcel E12=matrix(cp_all_percent)
putexcel F12=matrix(cp_all_se)
putexcel G12=matrix(cp_all_ll)
putexcel H12=matrix(cp_all_ul)
putexcel I12=matrix(mcp_all_percent)
putexcel J12=matrix(mcp_all_se)
putexcel K12=matrix(mcp_all_ll)
putexcel L12=matrix(mcp_all_ul)
putexcel M12=matrix(unmettot_all_percent)
putexcel N12=matrix(unmettot_all_se)
putexcel O12=matrix(unmettot_all_ll)
putexcel P12=matrix(unmettot_all_ul)
putexcel R12=matrix(cp_mar_percent)
putexcel S12=matrix(cp_mar_se)
putexcel T12=matrix(cp_mar_ll)
putexcel U12=matrix(cp_mar_ul)
putexcel V12=matrix(mcp_mar_percent)
putexcel W12=matrix(mcp_mar_se)
putexcel X12=matrix(mcp_mar_ll)
putexcel Y12=matrix(mcp_mar_ul)
putexcel Z12=matrix(unmettot_mar_percent)
putexcel AA12=matrix(unmettot_mar_se)
putexcel AB12=matrix(unmettot_mar_ll)
putexcel AC12=matrix(unmettot_mar_ul)


**** ROUND 6
clear
use "~/Dropbox (Gates Institute)/05_Ghana/PMAGH_Datasets/Round6/Prelim100/GHR6_WealthWeightAll_30Jul2018.dta", clear
putexcel B13=("Round 6")
putexcel C13=("7-9/2017")

** COUNT - Female Sample - All / Married Women  **
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & last_night==1
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel D13=matrix(FQresponse_1)
restore
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & last_night==1 & (FQmarital_status==1 | FQmarital_status==2)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel Q13=matrix(FQresponse_1)
restore 

** Estimate Percentage and 95% CI
keep if FRS_result==1 & HHQ_result==1 & last_night==1
egen all=tag(FQmetainstanceID)
egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)

svyset EA [pw=FQweight], strata(strata) singleunit(scaled)
foreach group in all mar {
preserve
	keep if `group'==1
	foreach indicator in cp mcp unmettot {
		svy: prop `indicator', citype(wilson)
		matrix reference=r(table)
		matrix `indicator'_`group'_percent=reference[1,2]*100	
		matrix `indicator'_`group'_se=reference[2,2]*100
		matrix `indicator'_`group'_ll=reference[5,2]*100
		matrix `indicator'_`group'_ul=reference[6,2]*100
		}	
	restore
	}
putexcel E13=matrix(cp_all_percent)
putexcel F13=matrix(cp_all_se)
putexcel G13=matrix(cp_all_ll)
putexcel H13=matrix(cp_all_ul)
putexcel I13=matrix(mcp_all_percent)
putexcel J13=matrix(mcp_all_se)
putexcel K13=matrix(mcp_all_ll)
putexcel L13=matrix(mcp_all_ul)
putexcel M13=matrix(unmettot_all_percent)
putexcel N13=matrix(unmettot_all_se)
putexcel O13=matrix(unmettot_all_ll)
putexcel P13=matrix(unmettot_all_ul)
putexcel R13=matrix(cp_mar_percent)
putexcel S13=matrix(cp_mar_se)
putexcel T13=matrix(cp_mar_ll)
putexcel U13=matrix(cp_mar_ul)
putexcel V13=matrix(mcp_mar_percent)
putexcel W13=matrix(mcp_mar_se)
putexcel X13=matrix(mcp_mar_ll)
putexcel Y13=matrix(mcp_mar_ul)
putexcel Z13=matrix(unmettot_mar_percent)
putexcel AA13=matrix(unmettot_mar_se)
putexcel AB13=matrix(unmettot_mar_ll)
putexcel AC13=matrix(unmettot_mar_ul)




***********************************************************************************
*** 		INDIA / RAJASTHAN
***********************************************************************************
capture clear all
local excel "RJ_KeyIndicators.xlsx"
local excel_sheet "CPR, mCPR, unmet need"
putexcel set "`excel'", modify sheet("`excel_sheet'")

***** ROUND 1
clear
use "~/Dropbox (Gates Institute)/06_India_Raj/PMARJ_Datasets/Round1/Final_PublicRelease/HHQ/PMA2016_INR1_Rajasthan_HHQFQ_v2_10Aug2018/PMA2016_INR1_Rajasthan_HHQFQ_v2_10Aug2018.dta", clear
putexcel A8=("PMA2020")
putexcel B8=("Round 1")
putexcel C8=("6-9/2016")

** COUNT - Female Sample - All / Married Women  **
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel D8=matrix(FQresponse_1)
restore
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3) & (FQmarital_status==1 | FQmarital_status==2)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel Q8=matrix(FQresponse_1)
restore 

*** Estimate Percentage and 95% CI
keep if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
egen all=tag(FQmetainstanceID)
egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)

svyset EA_ID [pw=FQweight], strata(strata) singleunit(scaled)
foreach group in all mar {
preserve
	keep if `group'==1
	foreach indicator in cp mcp unmettot {
		svy: prop `indicator', citype(wilson)
		matrix reference=r(table)
		matrix `indicator'_`group'_percent=reference[1,2]*100	
		matrix `indicator'_`group'_se=reference[2,2]*100
		matrix `indicator'_`group'_ll=reference[5,2]*100
		matrix `indicator'_`group'_ul=reference[6,2]*100
		}	
	restore
	}
run "/Users/asiewe/Documents/Data Analysis/DHS-PMA-Indicators/Descr_stats/Line8_stata2xcel.do"


***** ROUND 2
clear
use "~/Dropbox (Gates Institute)/06_India_Raj/PMARJ_Datasets/Round2/Final_PublicRelease/HHQ/PMA2017_INR2_Rajasthan_HHQFQ_v2_10Aug2018/PMA2017_INR2_Rajasthan_HHQFQ_v2_10Aug2018.dta", clear
putexcel A9=("PMA2020")
putexcel B9=("Round 2")
putexcel C9=("02-04/2017")

** COUNT - Female Sample - All / Married Women  **
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & last_night==1
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel D9=matrix(FQresponse_1)
restore
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & last_night==1 & (FQmarital_status==1 | FQmarital_status==2)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel Q9=matrix(FQresponse_1)
restore 

*** Estimate Percentage and 95% CI
keep if FRS_result==1 & HHQ_result==1 & last_night==1
egen all=tag(FQmetainstanceID)
egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)
svyset EA_ID [pw=FQweight], strata(strata) singleunit(scaled)
foreach group in all mar {
preserve
	keep if `group'==1
	foreach indicator in cp mcp unmettot {
		svy: prop `indicator', citype(wilson)
		matrix reference=r(table)
		matrix `indicator'_`group'_percent=reference[1,2]*100	
		matrix `indicator'_`group'_se=reference[2,2]*100
		matrix `indicator'_`group'_ll=reference[5,2]*100
		matrix `indicator'_`group'_ul=reference[6,2]*100
		}	
	restore
	}
run "/Users/asiewe/Documents/Data Analysis/DHS-PMA-Indicators/Descr_stats/Line9_stata2xcel.do"


***** ROUND 3
clear
use "~/Dropbox (Gates Institute)/06_India_Raj/PMARJ_Datasets/Round3/Final_PublicRelease/HHQ/PMA2017_INR3_Rajasthan_HHQFQ_v2_10Aug2018/PMA2017_INR3_Rajasthan_HHQFQ_v2_10Aug2018.dta", clear
putexcel A10=("PMA2020")
putexcel B10=("Round 3")
putexcel C10=("8-10/2017")

** COUNT - Female Sample - All / Married Women  **
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & last_night==1
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel D10=matrix(FQresponse_1)
restore
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & last_night==1 & (FQmarital_status==1 | FQmarital_status==2)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel Q10=matrix(FQresponse_1)
restore 

*** Estimate Percentage and 95% CI
keep if FRS_result==1 & HHQ_result==1 & last_night==1
egen all=tag(FQmetainstanceID)
egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)
svyset EA_ID [pw=FQweight], strata(strata) singleunit(scaled)
foreach group in all mar {
preserve
	keep if `group'==1
	foreach indicator in cp mcp unmettot {
		svy: prop `indicator', citype(wilson)
		matrix reference=r(table)
		matrix `indicator'_`group'_percent=reference[1,2]*100	
		matrix `indicator'_`group'_se=reference[2,2]*100
		matrix `indicator'_`group'_ll=reference[5,2]*100
		matrix `indicator'_`group'_ul=reference[6,2]*100
		}	
	restore
	}
run "/Users/asiewe/Documents/Data Analysis/DHS-PMA-Indicators/Descr_stats/Line10_stata2xcel.do"




***********************************************************************************
*** 		INDONESIA
***********************************************************************************
capture clear all
local excel "ID_KeyIndicators.xlsx"
local excel_sheet "CPR, mCPR, unmet need"
putexcel set "`excel'", modify sheet("`excel_sheet'")

***** ROUND 1
clear
use "~/Dropbox (Gates Institute)/07_Indonesia/PMAID_Datasets/Round1/Final_PublicRelease/HHQ/PMA2015_IDR1_HHQFQ_v2_10Aug2018/PMA2015_IDR1_HHQFQ_v2_10Aug2018.dta", clear
putexcel A8=("PMA2020")
putexcel B8=("Round 1")
putexcel C8=("5-8/2015")

** COUNT - Female Sample - All / Married Women  **
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel D8=matrix(FQresponse_1)
restore
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3) & (FQmarital_status==1 | FQmarital_status==2)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel Q8=matrix(FQresponse_1)
restore 

*** Estimate Percentage and 95% CI
keep if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
egen all=tag(FQmetainstanceID)
egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)
svyset EA_ID [pw=FQweight], strata(strata) singleunit(scaled)
foreach group in all mar {
preserve
	keep if `group'==1
	foreach indicator in cp mcp unmettot {
		svy: prop `indicator', citype(wilson)
		matrix reference=r(table)
		matrix `indicator'_`group'_percent=reference[1,2]*100	
		matrix `indicator'_`group'_se=reference[2,2]*100
		matrix `indicator'_`group'_ll=reference[5,2]*100
		matrix `indicator'_`group'_ul=reference[6,2]*100
		}	
	restore
	}
run "/Users/asiewe/Documents/Data Analysis/DHS-PMA-Indicators/Descr_stats/Line8_stata2xcel.do"


***** ROUND 2
clear
use "~/Dropbox (Gates Institute)/07_Indonesia/PMAID_Datasets/Round2/Final_PublicRelease/HHQ/PMA2016_IDR2_HHQFQ_v1_18Apr2018/PMA2016_IDR2_HHQFQ_v1_18Apr2018.dta", clear
putexcel A9=("PMA2020")
putexcel B9=("Round 2")
putexcel C9=("10/2016-01/2017")

** COUNT - Female Sample - All / Married Women  **
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & last_night==1
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel D9=matrix(FQresponse_1)
restore
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & last_night==1 & (FQmarital_status==1 | FQmarital_status==2)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel Q9=matrix(FQresponse_1)
restore 

*** Estimate Percentage and 95% CI
keep if FRS_result==1 & HHQ_result==1 & last_night==1
egen all=tag(FQmetainstanceID)
egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)
svyset EA_ID [pw=FQweight], strata(strata) singleunit(scaled)
foreach group in all mar {
preserve
	keep if `group'==1
	foreach indicator in cp mcp unmettot {
		svy: prop `indicator', citype(wilson)
		matrix reference=r(table)
		matrix `indicator'_`group'_percent=reference[1,2]*100	
		matrix `indicator'_`group'_se=reference[2,2]*100
		matrix `indicator'_`group'_ll=reference[5,2]*100
		matrix `indicator'_`group'_ul=reference[6,2]*100
		}	
	restore
	}
run "/Users/asiewe/Documents/Data Analysis/DHS-PMA-Indicators/Descr_stats/Line9_stata2xcel.do"



***********************************************************************************
*** 		KENYA
***********************************************************************************
capture clear all
local excel "KE_KeyIndicators.xlsx"
local excel_sheet "CPR, mCPR, unmet need"
putexcel set "`excel'", modify sheet("`excel_sheet'")

***** ROUND 1
clear
use "~/Dropbox (Gates Institute)/08_Kenya/PMAKE_Datasets/Round1/Final_PublicRelease/HHQ/PMA2014_KER1_HHQFQ_v4_13Aug2018/PMA2014_KER1_HHQFQ_v4_13Aug2018.dta", clear
putexcel A8=("PMA2020")
putexcel B8=("Round 1")
putexcel C8=("5-7/2014")

** COUNT - Female Sample - All / Married Women  **
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel D8=matrix(FQresponse_1)
restore
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3) & (FQmarital_status==1 | FQmarital_status==2)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel Q8=matrix(FQresponse_1)
restore 

*** Estimate Percentage and 95% CI
keep if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
egen all=tag(FQmetainstanceID)
egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)

svyset EA_ID [pw=FQweight], strata(strata) singleunit(scaled)
foreach group in all mar {
preserve
	keep if `group'==1
	foreach indicator in cp mcp unmettot {
		svy: prop `indicator', citype(wilson)
		matrix reference=r(table)
		matrix `indicator'_`group'_percent=reference[1,2]*100	
		matrix `indicator'_`group'_se=reference[2,2]*100
		matrix `indicator'_`group'_ll=reference[5,2]*100
		matrix `indicator'_`group'_ul=reference[6,2]*100
		}	
	restore
	}
run "/Users/asiewe/Documents/Data Analysis/DHS-PMA-Indicators/Descr_stats/Line8_stata2xcel.do"


***** ROUND 2
clear
use "~/Dropbox (Gates Institute)/08_Kenya/PMAKE_Datasets/Round2/Final_PublicRelease/HHQ/PMA2014_KER2_HHQFQ_v2_13Aug2018/PMA2014_KER2_HHQFQ_v2_13Aug2018.dta", clear
putexcel A9=("PMA2020")
putexcel B9=("Round 2")
putexcel C9=("11-12/2014")

** COUNT - Female Sample - All / Married Women  **
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel D9=matrix(FQresponse_1)
restore
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3) & (FQmarital_status==1 | FQmarital_status==2)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel Q9=matrix(FQresponse_1)
restore 

*** Estimate Percentage and 95% CI
keep if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
egen all=tag(FQmetainstanceID)
egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)
svyset EA_ID [pw=FQweight], strata(strata) singleunit(scaled)
foreach group in all mar {
preserve
	keep if `group'==1
	foreach indicator in cp mcp unmettot {
		svy: prop `indicator', citype(wilson)
		matrix reference=r(table)
		matrix `indicator'_`group'_percent=reference[1,2]*100	
		matrix `indicator'_`group'_se=reference[2,2]*100
		matrix `indicator'_`group'_ll=reference[5,2]*100
		matrix `indicator'_`group'_ul=reference[6,2]*100
		}	
	restore
	}
run "/Users/asiewe/Documents/Data Analysis/DHS-PMA-Indicators/Descr_stats/Line9_stata2xcel.do"


***** ROUND 3
clear
use "~/Dropbox (Gates Institute)/08_Kenya/PMAKE_Datasets/Round3/Final_PublicRelease/HHQ/PMA2015_KER3_HHQFQ_v2_13Aug2018/PMA2015_KER3_HHQFQ_v2_13Aug2018.dta", clear
putexcel A10=("PMA2020")
putexcel B10=("Round 3")
putexcel C10=("6-7/2015")

** COUNT - Female Sample - All / Married Women  **
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel D10=matrix(FQresponse_1)
restore
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3) & (FQmarital_status==1 | FQmarital_status==2)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel Q10=matrix(FQresponse_1)
restore 

*** Estimate Percentage and 95% CI
keep if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
egen all=tag(FQmetainstanceID)
egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)
svyset EA_ID [pw=FQweight], strata(strata) singleunit(scaled)
foreach group in all mar {
preserve
	keep if `group'==1
	foreach indicator in cp mcp unmettot {
		svy: prop `indicator', citype(wilson)
		matrix reference=r(table)
		matrix `indicator'_`group'_percent=reference[1,2]*100	
		matrix `indicator'_`group'_se=reference[2,2]*100
		matrix `indicator'_`group'_ll=reference[5,2]*100
		matrix `indicator'_`group'_ul=reference[6,2]*100
		}	
	restore
	}
run "/Users/asiewe/Documents/Data Analysis/DHS-PMA-Indicators/Descr_stats/Line10_stata2xcel.do"



***** ROUND 4
clear
use "~/Dropbox (Gates Institute)/08_Kenya/PMAKE_Datasets/Round4/Final_PublicRelease/HHQ/PMA2015_KER4_HHQFQ_v2_13Aug2018/PMA2015_KER4_HHQFQ_v2_13Aug2018.dta", clear
putexcel A11=("PMA2020")
putexcel B11=("Round 4")
putexcel C11=("11-12/2015")

** COUNT - Female Sample - All / Married Women  **
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel D11=matrix(FQresponse_1)
restore
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3) & (FQmarital_status==1 | FQmarital_status==2)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel Q11=matrix(FQresponse_1)
restore 

*** Estimate Percentage and 95% CI
keep if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
egen all=tag(FQmetainstanceID)
egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)
svyset EA_ID [pw=FQweight], strata(strata) singleunit(scaled)
foreach group in all mar {
preserve
	keep if `group'==1
	foreach indicator in cp mcp unmettot {
		svy: prop `indicator', citype(wilson)
		matrix reference=r(table)
		matrix `indicator'_`group'_percent=reference[1,2]*100	
		matrix `indicator'_`group'_se=reference[2,2]*100
		matrix `indicator'_`group'_ll=reference[5,2]*100
		matrix `indicator'_`group'_ul=reference[6,2]*100
		}	
	restore
	}
run "/Users/asiewe/Documents/Data Analysis/DHS-PMA-Indicators/Descr_stats/Line11_stata2xcel.do"


***** ROUND 5
clear
use "~/Dropbox (Gates Institute)/08_Kenya/PMAKE_Datasets/Round5/Final_PublicRelease/HHQ/PMA2016_KER5_HHQFQ_v2_13Aug2018/PMA2016_KER5_HHQFQ_v2_13Aug2018.dta", clear
putexcel A12=("PMA2020")
putexcel B12=("Round 5")
putexcel C12=("11-12/2016")

** COUNT - Female Sample - All / Married Women  **
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & last_night==1
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel D12=matrix(FQresponse_1)
restore
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & last_night==1 & (FQmarital_status==1 | FQmarital_status==2)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel Q12=matrix(FQresponse_1)
restore 

*** Estimate Percentage and 95% CI
keep if FRS_result==1 & HHQ_result==1 & last_night==1
egen all=tag(FQmetainstanceID)
egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)
svyset EA_ID [pw=FQweight], strata(strata) singleunit(scaled)
foreach group in all mar {
preserve
	keep if `group'==1
	foreach indicator in cp mcp unmettot {
		svy: prop `indicator', citype(wilson)
		matrix reference=r(table)
		matrix `indicator'_`group'_percent=reference[1,2]*100	
		matrix `indicator'_`group'_se=reference[2,2]*100
		matrix `indicator'_`group'_ll=reference[5,2]*100
		matrix `indicator'_`group'_ul=reference[6,2]*100
		}	
	restore
	}
run "/Users/asiewe/Documents/Data Analysis/DHS-PMA-Indicators/Descr_stats/Line12_stata2xcel.do"


***** ROUND 6
clear
use "~/Dropbox (Gates Institute)/08_Kenya/PMAKE_Datasets/Round6/Prelim100/KER6_WealthWeightAll_1Mar2018.dta", clear
putexcel A13=("PMA2020")
putexcel B13=("Round 6")
putexcel C13=("11-12/2017")

** COUNT - Female Sample - All / Married Women  **
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & last_night==1
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel D13=matrix(FQresponse_1)
restore
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & last_night==1 & (FQmarital_status==1 | FQmarital_status==2)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel Q13=matrix(FQresponse_1)
restore 

*** Estimate Percentage and 95% CI
keep if FRS_result==1 & HHQ_result==1 & last_night==1
egen all=tag(FQmetainstanceID)
egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)
svyset EA [pw=FQweight], strata(strata) singleunit(scaled)
foreach group in all mar {
preserve
	keep if `group'==1
	foreach indicator in cp mcp unmettot {
		svy: prop `indicator', citype(wilson)
		matrix reference=r(table)
		matrix `indicator'_`group'_percent=reference[1,2]*100	
		matrix `indicator'_`group'_se=reference[2,2]*100
		matrix `indicator'_`group'_ll=reference[5,2]*100
		matrix `indicator'_`group'_ul=reference[6,2]*100
		}	
	restore
	}
run "/Users/asiewe/Documents/Data Analysis/DHS-PMA-Indicators/Descr_stats/Line13_stata2xcel.do"




***********************************************************************************
*** 		NIGER (NATIONAL & NIAMEY)
***********************************************************************************
local excel "NE_KeyIndicators.xlsx"
local excel_sheet "CPR, mCPR, unmet need"
putexcel set "`excel'", modify sheet("`excel_sheet'")

***** ROUND 1 - NIAMEY 
clear
use "~/Dropbox (Gates Institute)/09_Niger/PMANE_Datasets/Round1/Final_PublicRelease/HHQ/PMA2015_NER1_Niamey_HHQFQ_v4_10Aug2018/PMA2015_NER1_Niamey_HHQFQ_v4_10Aug2018.dta", clear
putexcel A8=("PMA2020-Niamey")
putexcel B8=("Round 1")
putexcel C8=("6-8/2015")

** COUNT - Female Sample - All / Married Women  **
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel D8=matrix(FQresponse_1)
restore
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3) & (FQmarital_status==1 | FQmarital_status==2)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel Q8=matrix(FQresponse_1)
restore 

*** Estimate Percentage and 95% CI
keep if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
egen all=tag(FQmetainstanceID)
egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)
svyset EA_ID [pw=FQweight], strata(strata) singleunit(scaled)
foreach group in all mar {
preserve
	keep if `group'==1
	foreach indicator in cp mcp unmettot {
		svy: prop `indicator', citype(wilson)
		matrix reference=r(table)
		matrix `indicator'_`group'_percent=reference[1,2]*100	
		matrix `indicator'_`group'_se=reference[2,2]*100
		matrix `indicator'_`group'_ll=reference[5,2]*100
		matrix `indicator'_`group'_ul=reference[6,2]*100
		}	
	restore
	}
run "/Users/asiewe/Documents/Data Analysis/DHS-PMA-Indicators/Descr_stats/Line8_stata2xcel.do"


***** ROUND 2
clear
use "~/Dropbox (Gates Institute)/09_Niger/PMANE_Datasets/Round2/Final_PublicRelease/HHQ/PMA2016_NER2_National_HHQFQ_v2_10Aug2018/PMA2016_NER2_National_HHQFQ_v2_10Aug2018.dta", clear
putexcel A9=("PMA2020-National")
putexcel B9=("Round 2")
putexcel C9=("3-5/2016")

** COUNT - Female Sample - All / Married Women  **
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel D9=matrix(FQresponse_1)
restore
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3) & (FQmarital_status==1 | FQmarital_status==2)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel Q9=matrix(FQresponse_1)
restore 

*** Estimate Percentage and 95% CI
keep if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
egen all=tag(FQmetainstanceID)
egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)
svyset EA_ID [pw=FQweight], strata(strata) singleunit(scaled)
foreach group in all mar {
preserve
	keep if `group'==1
	foreach indicator in cp mcp unmettot {
		svy: prop `indicator', citype(wilson)
		matrix reference=r(table)
		matrix `indicator'_`group'_percent=reference[1,2]*100	
		matrix `indicator'_`group'_se=reference[2,2]*100
		matrix `indicator'_`group'_ll=reference[5,2]*100
		matrix `indicator'_`group'_ul=reference[6,2]*100
		}	
	restore
	}
run "/Users/asiewe/Documents/Data Analysis/DHS-PMA-Indicators/Descr_stats/Line9_stata2xcel.do"



***** ROUND 3 - NIAMEY 
clear
use "~/Dropbox (Gates Institute)/09_Niger/PMANE_Datasets/Round3/Final_PublicRelease/HHQ/PMA2016_NER3_Niamey_HHQFQ_v3_10Aug2018/PMA2016_NER3_Niamey_HHQFQ_v3_10Aug2018.dta", clear
putexcel A10=("PMA2020-Niamey")
putexcel B10=("Round 3")
putexcel C10=("11-12/2016")

** COUNT - Female Sample - All / Married Women  **
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & last_night==1
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel D10=matrix(FQresponse_1)
restore
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & last_night==1 & (FQmarital_status==1 | FQmarital_status==2)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel Q10=matrix(FQresponse_1)
restore 

*** Estimate Percentage and 95% CI
keep if FRS_result==1 & HHQ_result==1 & last_night==1
egen all=tag(FQmetainstanceID)
egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)
svyset EA_ID [pw=FQweight], strata(strata) singleunit(scaled)
foreach group in all mar {
preserve
	keep if `group'==1
	foreach indicator in cp mcp unmettot {
		svy: prop `indicator', citype(wilson)
		matrix reference=r(table)
		matrix `indicator'_`group'_percent=reference[1,2]*100	
		matrix `indicator'_`group'_se=reference[2,2]*100
		matrix `indicator'_`group'_ll=reference[5,2]*100
		matrix `indicator'_`group'_ul=reference[6,2]*100
		}	
	restore
	}
run "/Users/asiewe/Documents/Data Analysis/DHS-PMA-Indicators/Descr_stats/Line10_stata2xcel.do"

***** ROUND 4
clear
use "~/Dropbox (Gates Institute)/09_Niger/PMANE_Datasets/Round4/Final_PublicRelease/HHQ/PMA2017_NER4_National_HHQFQ_v2_10Aug2018/PMA2017_NER4_National_HHQFQ_v2_10Aug2018.dta", clear
putexcel A11=("PMA2020-National")
putexcel B11=("Round 4")
putexcel C11=("5-9/2017")

** COUNT - Female Sample - All / Married Women  **
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & last_night==1
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel D11=matrix(FQresponse_1)
restore
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & last_night==1 & (FQmarital_status==1 | FQmarital_status==2)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel Q11=matrix(FQresponse_1)
restore 

*** Estimate Percentage and 95% CI
keep if FRS_result==1 & HHQ_result==1 & last_night==1
egen all=tag(FQmetainstanceID)
egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)
svyset EA_ID [pw=FQweight], strata(strata) singleunit(scaled)
foreach group in all mar {
preserve
	keep if `group'==1
	foreach indicator in cp mcp unmettot {
		svy: prop `indicator', citype(wilson)
		matrix reference=r(table)
		matrix `indicator'_`group'_percent=reference[1,2]*100	
		matrix `indicator'_`group'_se=reference[2,2]*100
		matrix `indicator'_`group'_ll=reference[5,2]*100
		matrix `indicator'_`group'_ul=reference[6,2]*100
		}	
	restore
	}
run "/Users/asiewe/Documents/Data Analysis/DHS-PMA-Indicators/Descr_stats/Line11_stata2xcel.do"



***** ROUND 5 (ONGOING)
*/



***********************************************************************************
*** 		NIGERIA
***********************************************************************************
capture clear all
local excel "NG_KeyIndicators.xlsx"
local excel_sheet "CPR, mCPR, unmet need"
putexcel set "`excel'", modify sheet("`excel_sheet'")

***** ROUND 1 - LAGOS 
clear
*use "~/Dropbox (Gates Institute)/10_Nigeria/PMANG_Datasets/Round1/Prelim100/NGR1Lagos_WealthWeightAll_27Oct2016.dta", clear
use "~/Dropbox (Gates Institute)/10_Nigeria/PMANG_Datasets/Round1/Final_PublicRelease/HHQ/PMA2014_NGR1_Kaduna_Lagos_HHQFQ_v2_16Aug2018/PMA2014_NGR1_Lagos_HHQFQ_v2_16Aug2018.dta", clear
putexcel A8=("PMA2020-Lagos")
putexcel B8=("Round 1")
putexcel C8=("9-10/2014")
*tab1 state strata ur

** COUNT - Female Sample - All / Married Women  **
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel D8=matrix(FQresponse_1)
restore
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3) & (FQmarital_status==1 | FQmarital_status==2)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel Q8=matrix(FQresponse_1)
restore 

*** Estimate Percentage and 95% CI
keep if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
egen all=tag(FQmetainstanceID)
egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)
svyset Cluster_ID [pw=FQweight_Lagos], singleunit(scaled)
foreach group in all mar {
preserve
	keep if `group'==1
	foreach indicator in cp mcp unmettot {
		svy: prop `indicator', citype(wilson)
		matrix reference=r(table)
		matrix `indicator'_`group'_percent=reference[1,2]*100	
		matrix `indicator'_`group'_se=reference[2,2]*100
		matrix `indicator'_`group'_ll=reference[5,2]*100
		matrix `indicator'_`group'_ul=reference[6,2]*100
		}	
	restore
	}
run "/Users/asiewe/Documents/Data Analysis/DHS-PMA-Indicators/Descr_stats/Line8_stata2xcel.do"


***** ROUND 1 - KADUNA
clear
use "~/Dropbox (Gates Institute)/10_Nigeria/PMANG_Datasets/Round1/Final_PublicRelease/HHQ/PMA2014_NGR1_Kaduna_Lagos_HHQFQ_v2_16Aug2018/PMA2014_NGR1_Kaduna_HHQFQ_v2_16Aug2018.dta", clear
putexcel A9=("PMA2020-Kaduna")
putexcel B9=("Round 1")
putexcel C9=("9-10/2014")

** COUNT - Female Sample - All / Married Women  **
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel D9=matrix(FQresponse_1)
restore
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3) & (FQmarital_status==1 | FQmarital_status==2)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel Q9=matrix(FQresponse_1)
restore 

*** Estimate Percentage and 95% CI
keep if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
egen all=tag(FQmetainstanceID)
egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)
svyset Cluster_ID [pw=FQweight_Kaduna], singleunit(scaled)
foreach group in all mar {
preserve
	keep if `group'==1
	foreach indicator in cp mcp unmettot {
		svy: prop `indicator', citype(wilson)
		matrix reference=r(table)
		matrix `indicator'_`group'_percent=reference[1,2]*100	
		matrix `indicator'_`group'_se=reference[2,2]*100
		matrix `indicator'_`group'_ll=reference[5,2]*100
		matrix `indicator'_`group'_ul=reference[6,2]*100
		}	
	restore
	}
run "/Users/asiewe/Documents/Data Analysis/DHS-PMA-Indicators/Descr_stats/Line9_stata2xcel.do"



***** ROUND 2 - LAGOS 
clear
use "~/Dropbox (Gates Institute)/10_Nigeria/PMANG_Datasets/Round2/Final_PublicRelease/HHQ/PMA2015_NGR2_Kaduna_Lagos_HHQFQ_v3_16Aug2018/PMA2015_NGR2_Lagos_HHQFQ_v3_16Aug2018.dta", clear
putexcel A10=("PMA2020-Lagos")
putexcel B10=("Round 2")
putexcel C10=("9-10/2015")
*tab1 state strata ur

** COUNT - Female Sample - All / Married Women  **
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel D10=matrix(FQresponse_1)
restore
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3) & (FQmarital_status==1 | FQmarital_status==2)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel Q10=matrix(FQresponse_1)
restore 

*** Estimate Percentage and 95% CI
keep if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
egen all=tag(FQmetainstanceID)
egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)
svyset Cluster_ID [pw=FQweight_Lagos], singleunit(scaled)
foreach group in all mar {
preserve
	keep if `group'==1
	foreach indicator in cp mcp unmettot {
		svy: prop `indicator', citype(wilson)
		matrix reference=r(table)
		matrix `indicator'_`group'_percent=reference[1,2]*100	
		matrix `indicator'_`group'_se=reference[2,2]*100
		matrix `indicator'_`group'_ll=reference[5,2]*100
		matrix `indicator'_`group'_ul=reference[6,2]*100
		}	
	restore
	}
run "/Users/asiewe/Documents/Data Analysis/DHS-PMA-Indicators/Descr_stats/Line10_stata2xcel.do"


***** ROUND 2 - KADUNA
clear
use "~/Dropbox (Gates Institute)/10_Nigeria/PMANG_Datasets/Round2/Final_PublicRelease/HHQ/PMA2015_NGR2_Kaduna_Lagos_HHQFQ_v3_16Aug2018/PMA2015_NGR2_Kaduna_HHQFQ_v3_16Aug2018.dta", clear
putexcel A11=("PMA2020-Kaduna")
putexcel B11=("Round 2")
putexcel C11=("9-10/2015")
tab1 state strata ur

** COUNT - Female Sample - All / Married Women  **
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel D11=matrix(FQresponse_1)
restore
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3) & (FQmarital_status==1 | FQmarital_status==2)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel Q11=matrix(FQresponse_1)
restore 

*** Estimate Percentage and 95% CI
keep if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
egen all=tag(FQmetainstanceID)
egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)
svyset Cluster_ID [pw=FQweight_Kaduna], strata(strata_kaduna) singleunit(scaled)
foreach group in all mar {
preserve
	keep if `group'==1
	foreach indicator in cp mcp unmettot {
		svy: prop `indicator', citype(wilson)
		matrix reference=r(table)
		matrix `indicator'_`group'_percent=reference[1,2]*100	
		matrix `indicator'_`group'_se=reference[2,2]*100
		matrix `indicator'_`group'_ll=reference[5,2]*100
		matrix `indicator'_`group'_ul=reference[6,2]*100
		}	
	restore
	}
run "/Users/asiewe/Documents/Data Analysis/DHS-PMA-Indicators/Descr_stats/Line11_stata2xcel.do"



***** ROUND 3 - NATIONAL
clear
use "~/Dropbox (Gates Institute)/10_Nigeria/PMANG_Datasets/Round3/Final_PublicRelease/HHQ/PMA2016_NGR3_National_HHQFQ_v4_16Aug2018/PMA2016_NGR3_National_HHQFQ_v4_16Aug2018.dta", clear
putexcel A12=("PMA2020-National")
putexcel B12=("Round 3")
putexcel C12=("5-6/2016")

** COUNT - Female Sample - All / Married Women  **
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel D12=matrix(FQresponse_1)
restore
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3) & (FQmarital_status==1 | FQmarital_status==2)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel Q12=matrix(FQresponse_1)
restore 

*** Estimate Percentage and 95% CI
keep if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
egen all=tag(FQmetainstanceID)
egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)
svyset Cluster_ID [pw=FQweight_National], strata(strata) singleunit(scaled)
foreach group in all mar {
preserve
	keep if `group'==1
	foreach indicator in cp mcp unmettot {
		svy: prop `indicator', citype(wilson)
		matrix reference=r(table)
		matrix `indicator'_`group'_percent=reference[1,2]*100	
		matrix `indicator'_`group'_se=reference[2,2]*100
		matrix `indicator'_`group'_ll=reference[5,2]*100
		matrix `indicator'_`group'_ul=reference[6,2]*100
		}	
	restore
	}
run "/Users/asiewe/Documents/Data Analysis/DHS-PMA-Indicators/Descr_stats/Line12_stata2xcel.do"


	***** ROUND 3 - LAGOS
	clear
	use "~/Dropbox (Gates Institute)/10_Nigeria/PMANG_Datasets/Round3/Final_PublicRelease/HHQ/PMA2016_NGR3_National_HHQFQ_v4_16Aug2018/PMA2016_NGR3_National_HHQFQ_v4_16Aug2018.dta", clear
	putexcel A13=("PMA2020-Lagos")
	putexcel B13=("Round 3")
	putexcel C13=("5-6/2016")
	keep if state==2
	tab1 state strata ur

	** COUNT - Female Sample - All / Married Women  **
	preserve
	gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
	collapse (count) FQresponse_1
	mkmat FQresponse_1
	putexcel D13=matrix(FQresponse_1)
	restore
	preserve
	gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3) & (FQmarital_status==1 | FQmarital_status==2)
	collapse (count) FQresponse_1
	mkmat FQresponse_1
	putexcel Q13=matrix(FQresponse_1)
	restore 

	*** Estimate Percentage and 95% CI
	keep if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
	egen all=tag(FQmetainstanceID)
	egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)
	svyset Cluster_ID [pw=FQweight_Lagos], singleunit(scaled)
	foreach group in all mar {
	preserve
		keep if `group'==1
		foreach indicator in cp mcp unmettot {
			svy: prop `indicator', citype(wilson)
			matrix reference=r(table)
			matrix `indicator'_`group'_percent=reference[1,2]*100	
			matrix `indicator'_`group'_se=reference[2,2]*100
			matrix `indicator'_`group'_ll=reference[5,2]*100
			matrix `indicator'_`group'_ul=reference[6,2]*100
			}	
		restore
		}
	run "/Users/asiewe/Documents/Data Analysis/DHS-PMA-Indicators/Descr_stats/Line13_stata2xcel.do"
	
	
	***** ROUND 3 - KADUNA
	clear
	use "~/Dropbox (Gates Institute)/10_Nigeria/PMANG_Datasets/Round3/Final_PublicRelease/HHQ/PMA2016_NGR3_National_HHQFQ_v4_16Aug2018/PMA2016_NGR3_National_HHQFQ_v4_16Aug2018.dta", clear
	putexcel A14=("PMA2020-Kaduna")
	putexcel B14=("Round 3")
	putexcel C14=("5-6/2016")
	keep if state==1
	tab1 state strata ur

	** COUNT - Female Sample - All / Married Women  **
	preserve
	gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
	collapse (count) FQresponse_1
	mkmat FQresponse_1
	putexcel D14=matrix(FQresponse_1)
	restore
	preserve
	gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3) & (FQmarital_status==1 | FQmarital_status==2)
	collapse (count) FQresponse_1
	mkmat FQresponse_1
	putexcel Q14=matrix(FQresponse_1)
	restore 

	*** Estimate Percentage and 95% CI
	keep if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
	egen all=tag(FQmetainstanceID)
	egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)
	svyset Cluster_ID [pw=FQweight_Kaduna], singleunit(scaled)
	foreach group in all mar {
	preserve
		keep if `group'==1
		foreach indicator in cp mcp unmettot {
			svy: prop `indicator', citype(wilson)
			matrix reference=r(table)
			matrix `indicator'_`group'_percent=reference[1,2]*100	
			matrix `indicator'_`group'_se=reference[2,2]*100
			matrix `indicator'_`group'_ll=reference[5,2]*100
			matrix `indicator'_`group'_ul=reference[6,2]*100
			}	
		restore
		}
	run "/Users/asiewe/Documents/Data Analysis/DHS-PMA-Indicators/Descr_stats/Line14_stata2xcel.do"

	
	***** ROUND 3 - TARABA
	clear
	use "~/Dropbox (Gates Institute)/10_Nigeria/PMANG_Datasets/Round3/Final_PublicRelease/HHQ/PMA2016_NGR3_National_HHQFQ_v4_16Aug2018/PMA2016_NGR3_National_HHQFQ_v4_16Aug2018.dta", clear
	putexcel A15=("PMA2020-Taraba")
	putexcel B15=("Round 3")
	putexcel C15=("5-6/2016")
	keep if state==3
	tab1 state strata ur

	** COUNT - Female Sample - All / Married Women  **
	preserve
	gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
	collapse (count) FQresponse_1
	mkmat FQresponse_1
	putexcel D15=matrix(FQresponse_1)
	restore
	preserve
	gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3) & (FQmarital_status==1 | FQmarital_status==2)
	collapse (count) FQresponse_1
	mkmat FQresponse_1
	putexcel Q15=matrix(FQresponse_1)
	restore 

	*** Estimate Percentage and 95% CI
	keep if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
	egen all=tag(FQmetainstanceID)
	egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)
	svyset Cluster_ID [pw=FQweight_Taraba], singleunit(scaled)
	foreach group in all mar {
	preserve
		keep if `group'==1
		foreach indicator in cp mcp unmettot {
			svy: prop `indicator', citype(wilson)
			matrix reference=r(table)
			matrix `indicator'_`group'_percent=reference[1,2]*100	
			matrix `indicator'_`group'_se=reference[2,2]*100
			matrix `indicator'_`group'_ll=reference[5,2]*100
			matrix `indicator'_`group'_ul=reference[6,2]*100
			}	
		restore
		}
	run "/Users/asiewe/Documents/Data Analysis/DHS-PMA-Indicators/Descr_stats/Line15_stata2xcel.do"
	
	***** ROUND 3 - KANO
	clear
	use "~/Dropbox (Gates Institute)/10_Nigeria/PMANG_Datasets/Round3/Final_PublicRelease/HHQ/PMA2016_NGR3_National_HHQFQ_v4_16Aug2018/PMA2016_NGR3_National_HHQFQ_v4_16Aug2018.dta", clear
	putexcel A16=("PMA2020-Kano")
	putexcel B16=("Round 3")
	putexcel C16=("5-6/2016")
	keep if state==4
	tab1 state strata ur

	** COUNT - Female Sample - All / Married Women  **
	preserve
	gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
	collapse (count) FQresponse_1
	mkmat FQresponse_1
	putexcel D16=matrix(FQresponse_1)
	restore
	preserve
	gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3) & (FQmarital_status==1 | FQmarital_status==2)
	collapse (count) FQresponse_1
	mkmat FQresponse_1
	putexcel Q16=matrix(FQresponse_1)
	restore 

	*** Estimate Percentage and 95% CI
	keep if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
	egen all=tag(FQmetainstanceID)
	egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)
	svyset Cluster_ID [pw=FQweight_Kano], singleunit(scaled)
	foreach group in all mar {
	preserve
		keep if `group'==1
		foreach indicator in cp mcp unmettot {
			svy: prop `indicator', citype(wilson)
			matrix reference=r(table)
			matrix `indicator'_`group'_percent=reference[1,2]*100	
			matrix `indicator'_`group'_se=reference[2,2]*100
			matrix `indicator'_`group'_ll=reference[5,2]*100
			matrix `indicator'_`group'_ul=reference[6,2]*100
			}	
		restore
		}
	run "/Users/asiewe/Documents/Data Analysis/DHS-PMA-Indicators/Descr_stats/Line16_stata2xcel.do"

	***** ROUND 3 - RIVERS
	clear
	use "~/Dropbox (Gates Institute)/10_Nigeria/PMANG_Datasets/Round3/Final_PublicRelease/HHQ/PMA2016_NGR3_National_HHQFQ_v4_16Aug2018/PMA2016_NGR3_National_HHQFQ_v4_16Aug2018.dta", clear
	putexcel A17=("PMA2020-Rivers")
	putexcel B17=("Round 3")
	putexcel C17=("5-6/2016")
	keep if state==5
	tab1 state strata ur

	** COUNT - Female Sample - All / Married Women  **
	preserve
	gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
	collapse (count) FQresponse_1
	mkmat FQresponse_1
	putexcel D17=matrix(FQresponse_1)
	restore
	preserve
	gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3) & (FQmarital_status==1 | FQmarital_status==2)
	collapse (count) FQresponse_1
	mkmat FQresponse_1
	putexcel Q17=matrix(FQresponse_1)
	restore 

	*** Estimate Percentage and 95% CI
	keep if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
	egen all=tag(FQmetainstanceID)
	egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)
	svyset Cluster_ID [pw=FQweight_Rivers], singleunit(scaled)
	foreach group in all mar {
	preserve
		keep if `group'==1
		foreach indicator in cp mcp unmettot {
			svy: prop `indicator', citype(wilson)
			matrix reference=r(table)
			matrix `indicator'_`group'_percent=reference[1,2]*100	
			matrix `indicator'_`group'_se=reference[2,2]*100
			matrix `indicator'_`group'_ll=reference[5,2]*100
			matrix `indicator'_`group'_ul=reference[6,2]*100
			}	
		restore
		}
	run "/Users/asiewe/Documents/Data Analysis/DHS-PMA-Indicators/Descr_stats/Line17_stata2xcel.do"
	
	***** ROUND 3 - NASARAWA
	clear
	use "~/Dropbox (Gates Institute)/10_Nigeria/PMANG_Datasets/Round3/Final_PublicRelease/HHQ/PMA2016_NGR3_National_HHQFQ_v4_16Aug2018/PMA2016_NGR3_National_HHQFQ_v4_16Aug2018.dta", clear
	putexcel A18=("PMA2020-Nasarawa")
	putexcel B18=("Round 3")
	putexcel C18=("5-6/2016")
	keep if state==6
	tab1 state strata ur

	** COUNT - Female Sample - All / Married Women  **
	preserve
	gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
	collapse (count) FQresponse_1
	mkmat FQresponse_1
	putexcel D18=matrix(FQresponse_1)
	restore
	preserve
	gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3) & (FQmarital_status==1 | FQmarital_status==2)
	collapse (count) FQresponse_1
	mkmat FQresponse_1
	putexcel Q18=matrix(FQresponse_1)
	restore 

	*** Estimate Percentage and 95% CI
	keep if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
	egen all=tag(FQmetainstanceID)
	egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)
	svyset Cluster_ID [pw=FQweight_Nasarawa], singleunit(scaled)
	foreach group in all mar {
	preserve
		keep if `group'==1
		foreach indicator in cp mcp unmettot {
			svy: prop `indicator', citype(wilson)
			matrix reference=r(table)
			matrix `indicator'_`group'_percent=reference[1,2]*100	
			matrix `indicator'_`group'_se=reference[2,2]*100
			matrix `indicator'_`group'_ll=reference[5,2]*100
			matrix `indicator'_`group'_ul=reference[6,2]*100
			}	
		restore
		}
	run "/Users/asiewe/Documents/Data Analysis/DHS-PMA-Indicators/Descr_stats/Line18_stata2xcel.do"

	***** ROUND 3 - ANAMBRA
	clear
	use "~/Dropbox (Gates Institute)/10_Nigeria/PMANG_Datasets/Round3/Final_PublicRelease/HHQ/PMA2016_NGR3_National_HHQFQ_v4_16Aug2019/PMA2016_NGR3_National_HHQFQ_v4_16Aug2019.dta", clear
	putexcel A19=("PMA2020-Anambra")
	putexcel B19=("Round 3")
	putexcel C19=("5-6/2016")
	keep if state==7
	tab1 state strata ur

	** COUNT - Female Sample - All / Married Women  **
	preserve
	gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
	collapse (count) FQresponse_1
	mkmat FQresponse_1
	putexcel D19=matrix(FQresponse_1)
	restore
	preserve
	gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3) & (FQmarital_status==1 | FQmarital_status==2)
	collapse (count) FQresponse_1
	mkmat FQresponse_1
	putexcel Q19=matrix(FQresponse_1)
	restore 

	*** Estimate Percentage and 95% CI
	keep if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
	egen all=tag(FQmetainstanceID)
	egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)
	svyset Cluster_ID [pw=FQweight_Anambra], singleunit(scaled)
	foreach group in all mar {
	preserve
		keep if `group'==1
		foreach indicator in cp mcp unmettot {
			svy: prop `indicator', citype(wilson)
			matrix reference=r(table)
			matrix `indicator'_`group'_percent=reference[1,2]*100	
			matrix `indicator'_`group'_se=reference[2,2]*100
			matrix `indicator'_`group'_ll=reference[5,2]*100
			matrix `indicator'_`group'_ul=reference[6,2]*100
			}	
		restore
		}
	run "/Users/asiewe/Documents/Data Analysis/DHS-PMA-Indicators/Descr_stats/Line19_stata2xcel.do"
	
	
***** ROUND 4 - NATIONAL
clear
use "~/Dropbox (Gates Institute)/10_Nigeria/PMANG_Datasets/Round4/Final_PublicRelease/HHQ/PMA2017_NGR4_National_HHQFQ_v2_16Aug2018/PMA2017_NGR4_National_HHQFQ_v2_16Aug2018.dta", clear
putexcel A20=("PMA2020-National")
putexcel B20=("Round 4")
putexcel C20=("3-4/2017")

** COUNT - Female Sample - All / Married Women  **
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & last_night==1
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel D20=matrix(FQresponse_1)
restore
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & last_night==1 & (FQmarital_status==1 | FQmarital_status==2)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel Q20=matrix(FQresponse_1)
restore 

*** Estimate Percentage and 95% CI
keep if FRS_result==1 & HHQ_result==1 & last_night==1
egen all=tag(FQmetainstanceID)
egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)
svyset Cluster_ID [pw=FQweight_National], strata(strata) singleunit(scaled)
foreach group in all mar {
preserve
	keep if `group'==1
	foreach indicator in cp mcp unmettot {
		svy: prop `indicator', citype(wilson)
		matrix reference=r(table)
		matrix `indicator'_`group'_percent=reference[1,2]*100	
		matrix `indicator'_`group'_se=reference[2,2]*100
		matrix `indicator'_`group'_ll=reference[5,2]*100
		matrix `indicator'_`group'_ul=reference[6,2]*100
		}	
	restore
	}
run "/Users/asiewe/Documents/Data Analysis/DHS-PMA-Indicators/Descr_stats/Line20_stata2xcel.do"


	***** ROUND 4 - LAGOS
	clear
	use "~/Dropbox (Gates Institute)/10_Nigeria/PMANG_Datasets/Round4/Final_PublicRelease/HHQ/PMA2017_NGR4_National_HHQFQ_v2_16Aug2018/PMA2017_NGR4_National_HHQFQ_v2_16Aug2018.dta", clear
	putexcel A21=("PMA2020-Lagos")
	putexcel B21=("Round 4")
	putexcel C21=("3-4/2017")
	keep if state==2
	tab1 state strata ur

	** COUNT - Female Sample - All / Married Women  **
	preserve
	gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & last_night==1
	collapse (count) FQresponse_1
	mkmat FQresponse_1
	putexcel D21=matrix(FQresponse_1)
	restore
	preserve
	gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & last_night==1 & (FQmarital_status==1 | FQmarital_status==2)
	collapse (count) FQresponse_1
	mkmat FQresponse_1
	putexcel Q21=matrix(FQresponse_1)
	restore 

	*** Estimate Percentage and 95% CI
	keep if FRS_result==1 & HHQ_result==1 & last_night==1
	egen all=tag(FQmetainstanceID)
	egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)
	svyset Cluster_ID [pw=FQweight_Lagos], singleunit(scaled)
	foreach group in all mar {
	preserve
		keep if `group'==1
		foreach indicator in cp mcp unmettot {
			svy: prop `indicator', citype(wilson)
			matrix reference=r(table)
			matrix `indicator'_`group'_percent=reference[1,2]*100	
			matrix `indicator'_`group'_se=reference[2,2]*100
			matrix `indicator'_`group'_ll=reference[5,2]*100
			matrix `indicator'_`group'_ul=reference[6,2]*100
			}	
		restore
		}
	run "/Users/asiewe/Documents/Data Analysis/DHS-PMA-Indicators/Descr_stats/Line21_stata2xcel.do"
	
	
	***** ROUND 4 - KADUNA
	clear
	use "~/Dropbox (Gates Institute)/10_Nigeria/PMANG_Datasets/Round4/Final_PublicRelease/HHQ/PMA2017_NGR4_National_HHQFQ_v2_16Aug2018/PMA2017_NGR4_National_HHQFQ_v2_16Aug2018.dta", clear
	putexcel A22=("PMA2020-Kaduna")
	putexcel B22=("Round 4")
	putexcel C22=("3-4/2017")
	keep if state==1
	tab1 state strata ur

	** COUNT - Female Sample - All / Married Women  **
	preserve
	gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & last_night==1
	collapse (count) FQresponse_1
	mkmat FQresponse_1
	putexcel D22=matrix(FQresponse_1)
	restore
	preserve
	gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & last_night==1 & (FQmarital_status==1 | FQmarital_status==2)
	collapse (count) FQresponse_1
	mkmat FQresponse_1
	putexcel Q22=matrix(FQresponse_1)
	restore 

	*** Estimate Percentage and 95% CI
	keep if FRS_result==1 & HHQ_result==1 & last_night==1
	egen all=tag(FQmetainstanceID)
	egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)
	svyset Cluster_ID [pw=FQweight_Kaduna], singleunit(scaled)
	foreach group in all mar {
	preserve
		keep if `group'==1
		foreach indicator in cp mcp unmettot {
			svy: prop `indicator', citype(wilson)
			matrix reference=r(table)
			matrix `indicator'_`group'_percent=reference[1,2]*100	
			matrix `indicator'_`group'_se=reference[2,2]*100
			matrix `indicator'_`group'_ll=reference[5,2]*100
			matrix `indicator'_`group'_ul=reference[6,2]*100
			}	
		restore
		}
	run "/Users/asiewe/Documents/Data Analysis/DHS-PMA-Indicators/Descr_stats/Line22_stata2xcel.do"

	
	***** ROUND 4 - TARABA
	clear
	use "~/Dropbox (Gates Institute)/10_Nigeria/PMANG_Datasets/Round4/Final_PublicRelease/HHQ/PMA2017_NGR4_National_HHQFQ_v2_16Aug2018/PMA2017_NGR4_National_HHQFQ_v2_16Aug2018.dta", clear
	putexcel A23=("PMA2020-Taraba")
	putexcel B23=("Round 4")
	putexcel C23=("3-4/2017")
	keep if state==3
	tab1 state strata ur

	** COUNT - Female Sample - All / Married Women  **
	preserve
	gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & last_night==1
	collapse (count) FQresponse_1
	mkmat FQresponse_1
	putexcel D23=matrix(FQresponse_1)
	restore
	preserve
	gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & last_night==1 & (FQmarital_status==1 | FQmarital_status==2)
	collapse (count) FQresponse_1
	mkmat FQresponse_1
	putexcel Q23=matrix(FQresponse_1)
	restore 

	*** Estimate Percentage and 95% CI
	keep if FRS_result==1 & HHQ_result==1 & last_night==1
	egen all=tag(FQmetainstanceID)
	egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)
	svyset Cluster_ID [pw=FQweight_Taraba], singleunit(scaled)
	foreach group in all mar {
	preserve
		keep if `group'==1
		foreach indicator in cp mcp unmettot {
			svy: prop `indicator', citype(wilson)
			matrix reference=r(table)
			matrix `indicator'_`group'_percent=reference[1,2]*100	
			matrix `indicator'_`group'_se=reference[2,2]*100
			matrix `indicator'_`group'_ll=reference[5,2]*100
			matrix `indicator'_`group'_ul=reference[6,2]*100
			}	
		restore
		}
	run "/Users/asiewe/Documents/Data Analysis/DHS-PMA-Indicators/Descr_stats/Line23_stata2xcel.do"
	
	***** ROUND 4 - KANO
	clear
	use "~/Dropbox (Gates Institute)/10_Nigeria/PMANG_Datasets/Round4/Final_PublicRelease/HHQ/PMA2017_NGR4_National_HHQFQ_v2_16Aug2018/PMA2017_NGR4_National_HHQFQ_v2_16Aug2018.dta", clear
	putexcel A24=("PMA2020-Kano")
	putexcel B24=("Round 4")
	putexcel C24=("3-4/2017")
	keep if state==4
	tab1 state strata ur

	** COUNT - Female Sample - All / Married Women  **
	preserve
	gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & last_night==1
	collapse (count) FQresponse_1
	mkmat FQresponse_1
	putexcel D24=matrix(FQresponse_1)
	restore
	preserve
	gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & last_night==1 & (FQmarital_status==1 | FQmarital_status==2)
	collapse (count) FQresponse_1
	mkmat FQresponse_1
	putexcel Q24=matrix(FQresponse_1)
	restore 

	*** Estimate Percentage and 95% CI
	keep if FRS_result==1 & HHQ_result==1 & last_night==1
	egen all=tag(FQmetainstanceID)
	egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)
	svyset Cluster_ID [pw=FQweight_Kano], singleunit(scaled)
	foreach group in all mar {
	preserve
		keep if `group'==1
		foreach indicator in cp mcp unmettot {
			svy: prop `indicator', citype(wilson)
			matrix reference=r(table)
			matrix `indicator'_`group'_percent=reference[1,2]*100	
			matrix `indicator'_`group'_se=reference[2,2]*100
			matrix `indicator'_`group'_ll=reference[5,2]*100
			matrix `indicator'_`group'_ul=reference[6,2]*100
			}	
		restore
		}
	run "/Users/asiewe/Documents/Data Analysis/DHS-PMA-Indicators/Descr_stats/Line24_stata2xcel.do"

	***** ROUND 4 - RIVERS
	clear
	use "~/Dropbox (Gates Institute)/10_Nigeria/PMANG_Datasets/Round4/Final_PublicRelease/HHQ/PMA2017_NGR4_National_HHQFQ_v2_16Aug2018/PMA2017_NGR4_National_HHQFQ_v2_16Aug2018.dta", clear
	putexcel A25=("PMA2020-Rivers")
	putexcel B25=("Round 4")
	putexcel C25=("3-4/2017")
	keep if state==5
	tab1 state strata ur

	** COUNT - Female Sample - All / Married Women  **
	preserve
	gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & last_night==1
	collapse (count) FQresponse_1
	mkmat FQresponse_1
	putexcel D25=matrix(FQresponse_1)
	restore
	preserve
	gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & last_night==1 & (FQmarital_status==1 | FQmarital_status==2)
	collapse (count) FQresponse_1
	mkmat FQresponse_1
	putexcel Q25=matrix(FQresponse_1)
	restore 

	*** Estimate Percentage and 95% CI
	keep if FRS_result==1 & HHQ_result==1 & last_night==1
	egen all=tag(FQmetainstanceID)
	egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)
	svyset Cluster_ID [pw=FQweight_Rivers], singleunit(scaled)
	foreach group in all mar {
	preserve
		keep if `group'==1
		foreach indicator in cp mcp unmettot {
			svy: prop `indicator', citype(wilson)
			matrix reference=r(table)
			matrix `indicator'_`group'_percent=reference[1,2]*100	
			matrix `indicator'_`group'_se=reference[2,2]*100
			matrix `indicator'_`group'_ll=reference[5,2]*100
			matrix `indicator'_`group'_ul=reference[6,2]*100
			}	
		restore
		}
	run "/Users/asiewe/Documents/Data Analysis/DHS-PMA-Indicators/Descr_stats/Line25_stata2xcel.do"
	
	***** ROUND 4 - NASARAWA
	clear
	use "~/Dropbox (Gates Institute)/10_Nigeria/PMANG_Datasets/Round4/Final_PublicRelease/HHQ/PMA2017_NGR4_National_HHQFQ_v2_16Aug2018/PMA2017_NGR4_National_HHQFQ_v2_16Aug2018.dta", clear
	putexcel A26=("PMA2020-Nasarawa")
	putexcel B26=("Round 4")
	putexcel C26=("3-4/2017")
	keep if state==6
	tab1 state strata ur

	** COUNT - Female Sample - All / Married Women  **
	preserve
	gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & last_night==1
	collapse (count) FQresponse_1
	mkmat FQresponse_1
	putexcel D26=matrix(FQresponse_1)
	restore
	preserve
	gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & last_night==1 & (FQmarital_status==1 | FQmarital_status==2)
	collapse (count) FQresponse_1
	mkmat FQresponse_1
	putexcel Q26=matrix(FQresponse_1)
	restore 

	*** Estimate Percentage and 95% CI
	keep if FRS_result==1 & HHQ_result==1 & last_night==1
	egen all=tag(FQmetainstanceID)
	egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)
	svyset Cluster_ID [pw=FQweight_Nasarawa], singleunit(scaled)
	foreach group in all mar {
	preserve
		keep if `group'==1
		foreach indicator in cp mcp unmettot {
			svy: prop `indicator', citype(wilson)
			matrix reference=r(table)
			matrix `indicator'_`group'_percent=reference[1,2]*100	
			matrix `indicator'_`group'_se=reference[2,2]*100
			matrix `indicator'_`group'_ll=reference[5,2]*100
			matrix `indicator'_`group'_ul=reference[6,2]*100
			}	
		restore
		}
	run "/Users/asiewe/Documents/Data Analysis/DHS-PMA-Indicators/Descr_stats/Line26_stata2xcel.do"

	***** ROUND 4 - ANAMBRA
	clear
	use "~/Dropbox (Gates Institute)/10_Nigeria/PMANG_Datasets/Round4/Final_PublicRelease/HHQ/PMA2016_NGR3_National_HHQFQ_v4_16Aug2019/PMA2016_NGR3_National_HHQFQ_v4_16Aug2019.dta", clear
	putexcel A27=("PMA2020-Anambra")
	putexcel B27=("Round 4")
	putexcel C27=("3-4/2017")
	keep if state==7
	tab1 state strata ur

	** COUNT - Female Sample - All / Married Women  **
	preserve
	gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & last_night==1
	collapse (count) FQresponse_1
	mkmat FQresponse_1
	putexcel D27=matrix(FQresponse_1)
	restore
	preserve
	gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & last_night==1 & (FQmarital_status==1 | FQmarital_status==2)
	collapse (count) FQresponse_1
	mkmat FQresponse_1
	putexcel Q27=matrix(FQresponse_1)
	restore 

	*** Estimate Percentage and 95% CI
	keep if FRS_result==1 & HHQ_result==1 & last_night==1
	egen all=tag(FQmetainstanceID)
	egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)
	svyset Cluster_ID [pw=FQweight_Anambra], singleunit(scaled)
	foreach group in all mar {
	preserve
		keep if `group'==1
		foreach indicator in cp mcp unmettot {
			svy: prop `indicator', citype(wilson)
			matrix reference=r(table)
			matrix `indicator'_`group'_percent=reference[1,2]*100	
			matrix `indicator'_`group'_se=reference[2,2]*100
			matrix `indicator'_`group'_ll=reference[5,2]*100
			matrix `indicator'_`group'_ul=reference[6,2]*100
			}	
		restore
		}
	run "/Users/asiewe/Documents/Data Analysis/DHS-PMA-Indicators/Descr_stats/Line27_stata2xcel.do"

	

***** ROUND 5
clear
use "~/Dropbox (Gates Institute)/10_Nigeria/PMANG_Datasets/Round5/Prelim100/HHQFQ/NGR5_WealthWeightAll_9Aug2018.dta", clear

*encoding state*
label define statenum_list 1 "Kaduna" 2 "Lagos" 3 "Taraba" 4 "Kano" 5 "Rivers" 6 "Nasarawa" 7 "Anambra"
encode State, gen(State2) lab(statenum_list)
order State2, after(State)
drop State
rename State2 state
rename FQweight FQweight_National

save "ngr5all", replace


***** NATIONAL
putexcel A28=("PMA2020-National")
putexcel B28=("Round 5")
putexcel C28=("11/2017-05/2018")

** COUNT - Female Sample - All / Married Women  **
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & last_night==1
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel D28=matrix(FQresponse_1)
restore
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & last_night==1 & (FQmarital_status==1 | FQmarital_status==2)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel Q28=matrix(FQresponse_1)
restore 

*** Estimate Percentage and 95% CI
keep if FRS_result==1 & HHQ_result==1 & last_night==1
egen all=tag(FQmetainstanceID)
egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)
svyset ClusterID [pw=FQweight_National], strata(strata) singleunit(scaled)
foreach group in all mar {
preserve
	keep if `group'==1
	foreach indicator in cp mcp unmettot {
		svy: prop `indicator', citype(wilson)
		matrix reference=r(table)
		matrix `indicator'_`group'_percent=reference[1,2]*100	
		matrix `indicator'_`group'_se=reference[2,2]*100
		matrix `indicator'_`group'_ll=reference[5,2]*100
		matrix `indicator'_`group'_ul=reference[6,2]*100
		}	
	restore
	}
run "/Users/asiewe/Documents/Data Analysis/DHS-PMA-Indicators/Descr_stats/Line28_stata2xcel.do"


	***** ROUND 5 - LAGOS
	clear
	use "~/Users/asiewe/Documents/Data Analysis/DHS-PMA-Indicators/Excel Output/ngr5all.dta", clear
	putexcel A29=("PMA2020-Lagos")
	putexcel B29=("Round 5")
	putexcel C29=("11/2017-05/2018")
	keep if state==2
	tab1 state strata ur

	** COUNT - Female Sample - All / Married Women  **
	preserve
	gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & last_night==1
	collapse (count) FQresponse_1
	mkmat FQresponse_1
	putexcel D29=matrix(FQresponse_1)
	restore
	preserve
	gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & last_night==1 & (FQmarital_status==1 | FQmarital_status==2)
	collapse (count) FQresponse_1
	mkmat FQresponse_1
	putexcel Q29=matrix(FQresponse_1)
	restore 

	*** Estimate Percentage and 95% CI
	keep if FRS_result==1 & HHQ_result==1 & last_night==1
	egen all=tag(FQmetainstanceID)
	egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)
	svyset Cluster_ID [pw=FQweight_Lagos], singleunit(scaled)
	foreach group in all mar {
	preserve
		keep if `group'==1
		foreach indicator in cp mcp unmettot {
			svy: prop `indicator', citype(wilson)
			matrix reference=r(table)
			matrix `indicator'_`group'_percent=reference[1,2]*100	
			matrix `indicator'_`group'_se=reference[2,2]*100
			matrix `indicator'_`group'_ll=reference[5,2]*100
			matrix `indicator'_`group'_ul=reference[6,2]*100
			}	
		restore
		}
	run "/Users/asiewe/Documents/Data Analysis/DHS-PMA-Indicators/Descr_stats/Line29_stata2xcel.do"
	
	
	***** ROUND 5 - KADUNA
	clear
	use "~/Users/asiewe/Documents/Data Analysis/DHS-PMA-Indicators/Excel Output/ngr5all.dta", clear
	putexcel A30=("PMA2020-Kaduna")
	putexcel B30=("Round 5")
	putexcel C30=("11/2017-05/2018")
	keep if state==1
	tab1 state strata ur

	** COUNT - Female Sample - All / Married Women  **
	preserve
	gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & last_night==1
	collapse (count) FQresponse_1
	mkmat FQresponse_1
	putexcel D30=matrix(FQresponse_1)
	restore
	preserve
	gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & last_night==1 & (FQmarital_status==1 | FQmarital_status==2)
	collapse (count) FQresponse_1
	mkmat FQresponse_1
	putexcel Q30=matrix(FQresponse_1)
	restore 

	*** Estimate Percentage and 95% CI
	keep if FRS_result==1 & HHQ_result==1 & last_night==1
	egen all=tag(FQmetainstanceID)
	egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)
	svyset Cluster_ID [pw=FQweight_Kaduna], singleunit(scaled)
	foreach group in all mar {
	preserve
		keep if `group'==1
		foreach indicator in cp mcp unmettot {
			svy: prop `indicator', citype(wilson)
			matrix reference=r(table)
			matrix `indicator'_`group'_percent=reference[1,2]*100	
			matrix `indicator'_`group'_se=reference[2,2]*100
			matrix `indicator'_`group'_ll=reference[5,2]*100
			matrix `indicator'_`group'_ul=reference[6,2]*100
			}	
		restore
		}
	run "/Users/asiewe/Documents/Data Analysis/DHS-PMA-Indicators/Descr_stats/Line30_stata2xcel.do"

	
	***** ROUND 5 - TARABA
	clear
	use "~/Users/asiewe/Documents/Data Analysis/DHS-PMA-Indicators/Excel Output/ngr5all.dta", clear
	putexcel A31=("PMA2020-Taraba")
	putexcel B31=("Round 5")
	putexcel C31=("11/2017-05/2018")
	keep if state==3
	tab1 state strata ur

	** COUNT - Female Sample - All / Married Women  **
	preserve
	gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & last_night==1
	collapse (count) FQresponse_1
	mkmat FQresponse_1
	putexcel D31=matrix(FQresponse_1)
	restore
	preserve
	gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & last_night==1 & (FQmarital_status==1 | FQmarital_status==2)
	collapse (count) FQresponse_1
	mkmat FQresponse_1
	putexcel Q31=matrix(FQresponse_1)
	restore 

	*** Estimate Percentage and 95% CI
	keep if FRS_result==1 & HHQ_result==1 & last_night==1
	egen all=tag(FQmetainstanceID)
	egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)
	svyset Cluster_ID [pw=FQweight_Taraba], singleunit(scaled)
	foreach group in all mar {
	preserve
		keep if `group'==1
		foreach indicator in cp mcp unmettot {
			svy: prop `indicator', citype(wilson)
			matrix reference=r(table)
			matrix `indicator'_`group'_percent=reference[1,2]*100	
			matrix `indicator'_`group'_se=reference[2,2]*100
			matrix `indicator'_`group'_ll=reference[5,2]*100
			matrix `indicator'_`group'_ul=reference[6,2]*100
			}	
		restore
		}
	run "/Users/asiewe/Documents/Data Analysis/DHS-PMA-Indicators/Descr_stats/Line31_stata2xcel.do"
	
	***** ROUND 5 - KANO
	clear
	use "~/Users/asiewe/Documents/Data Analysis/DHS-PMA-Indicators/Excel Output/ngr5all.dta", clear
	putexcel A32=("PMA2020-Kano")
	putexcel B32=("Round 5")
	putexcel C32=("11/2017-05/2018")
	keep if state==4
	tab1 state strata ur

	** COUNT - Female Sample - All / Married Women  **
	preserve
	gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & last_night==1
	collapse (count) FQresponse_1
	mkmat FQresponse_1
	putexcel D32=matrix(FQresponse_1)
	restore
	preserve
	gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & last_night==1 & (FQmarital_status==1 | FQmarital_status==2)
	collapse (count) FQresponse_1
	mkmat FQresponse_1
	putexcel Q32=matrix(FQresponse_1)
	restore 

	*** Estimate Percentage and 95% CI
	keep if FRS_result==1 & HHQ_result==1 & last_night==1
	egen all=tag(FQmetainstanceID)
	egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)
	svyset Cluster_ID [pw=FQweight_Kano], singleunit(scaled)
	foreach group in all mar {
	preserve
		keep if `group'==1
		foreach indicator in cp mcp unmettot {
			svy: prop `indicator', citype(wilson)
			matrix reference=r(table)
			matrix `indicator'_`group'_percent=reference[1,2]*100	
			matrix `indicator'_`group'_se=reference[2,2]*100
			matrix `indicator'_`group'_ll=reference[5,2]*100
			matrix `indicator'_`group'_ul=reference[6,2]*100
			}	
		restore
		}
	run "/Users/asiewe/Documents/Data Analysis/DHS-PMA-Indicators/Descr_stats/Line32_stata2xcel.do"

	***** ROUND 5 - RIVERS
	clear
	use "~/Users/asiewe/Documents/Data Analysis/DHS-PMA-Indicators/Excel Output/ngr5all.dta", clear
	putexcel A33=("PMA2020-Rivers")
	putexcel B33=("Round 5")
	putexcel C33=("11/2017-05/2018")
	keep if state==5
	tab1 state strata ur

	** COUNT - Female Sample - All / Married Women  **
	preserve
	gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & last_night==1
	collapse (count) FQresponse_1
	mkmat FQresponse_1
	putexcel D33=matrix(FQresponse_1)
	restore
	preserve
	gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & last_night==1 & (FQmarital_status==1 | FQmarital_status==2)
	collapse (count) FQresponse_1
	mkmat FQresponse_1
	putexcel Q33=matrix(FQresponse_1)
	restore 

	*** Estimate Percentage and 95% CI
	keep if FRS_result==1 & HHQ_result==1 & last_night==1
	egen all=tag(FQmetainstanceID)
	egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)
	svyset Cluster_ID [pw=FQweight_Rivers], singleunit(scaled)
	foreach group in all mar {
	preserve
		keep if `group'==1
		foreach indicator in cp mcp unmettot {
			svy: prop `indicator', citype(wilson)
			matrix reference=r(table)
			matrix `indicator'_`group'_percent=reference[1,2]*100	
			matrix `indicator'_`group'_se=reference[2,2]*100
			matrix `indicator'_`group'_ll=reference[5,2]*100
			matrix `indicator'_`group'_ul=reference[6,2]*100
			}	
		restore
		}
	run "/Users/asiewe/Documents/Data Analysis/DHS-PMA-Indicators/Descr_stats/Line33_stata2xcel.do"
	
	***** ROUND 5 - NASARAWA
	clear
	use "~/Users/asiewe/Documents/Data Analysis/DHS-PMA-Indicators/Excel Output/ngr5all.dta", clear
	putexcel A34=("PMA2020-Nasarawa")
	putexcel B34=("Round 5")
	putexcel C34=("11/2017-05/2018")
	keep if state==6
	tab1 state strata ur

	** COUNT - Female Sample - All / Married Women  **
	preserve
	gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & last_night==1
	collapse (count) FQresponse_1
	mkmat FQresponse_1
	putexcel D34=matrix(FQresponse_1)
	restore
	preserve
	gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & last_night==1 & (FQmarital_status==1 | FQmarital_status==2)
	collapse (count) FQresponse_1
	mkmat FQresponse_1
	putexcel Q34=matrix(FQresponse_1)
	restore 

	*** Estimate Percentage and 95% CI
	keep if FRS_result==1 & HHQ_result==1 & last_night==1
	egen all=tag(FQmetainstanceID)
	egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)
	svyset Cluster_ID [pw=FQweight_Nasarawa], singleunit(scaled)
	foreach group in all mar {
	preserve
		keep if `group'==1
		foreach indicator in cp mcp unmettot {
			svy: prop `indicator', citype(wilson)
			matrix reference=r(table)
			matrix `indicator'_`group'_percent=reference[1,2]*100	
			matrix `indicator'_`group'_se=reference[2,2]*100
			matrix `indicator'_`group'_ll=reference[5,2]*100
			matrix `indicator'_`group'_ul=reference[6,2]*100
			}	
		restore
		}
	run "/Users/asiewe/Documents/Data Analysis/DHS-PMA-Indicators/Descr_stats/Line34_stata2xcel.do"

	***** ROUND 5 - ANAMBRA
	clear
	use "~/Dropbox (Gates Institute)/10_Nigeria/PMANG_Datasets/Round4/Final_PublicRelease/HHQ/PMA2016_NGR3_National_HHQFQ_v4_16Aug2019/PMA2016_NGR3_National_HHQFQ_v4_16Aug2019.dta", clear
	putexcel A35=("PMA2020-Anambra")
	putexcel B35=("Round 5")
	putexcel C35=("11/2017-05/2018")
	keep if state==7
	tab1 state strata ur

	** COUNT - Female Sample - All / Married Women  **
	preserve
	gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & last_night==1
	collapse (count) FQresponse_1
	mkmat FQresponse_1
	putexcel D35=matrix(FQresponse_1)
	restore
	preserve
	gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & last_night==1 & (FQmarital_status==1 | FQmarital_status==2)
	collapse (count) FQresponse_1
	mkmat FQresponse_1
	putexcel Q35=matrix(FQresponse_1)
	restore 

	*** Estimate Percentage and 95% CI
	keep if FRS_result==1 & HHQ_result==1 & last_night==1
	egen all=tag(FQmetainstanceID)
	egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)
	svyset Cluster_ID [pw=FQweight_Anambra], singleunit(scaled)
	foreach group in all mar {
	preserve
		keep if `group'==1
		foreach indicator in cp mcp unmettot {
			svy: prop `indicator', citype(wilson)
			matrix reference=r(table)
			matrix `indicator'_`group'_percent=reference[1,2]*100	
			matrix `indicator'_`group'_se=reference[2,2]*100
			matrix `indicator'_`group'_ll=reference[5,2]*100
			matrix `indicator'_`group'_ul=reference[6,2]*100
			}	
		restore
		}
	run "/Users/asiewe/Documents/Data Analysis/DHS-PMA-Indicators/Descr_stats/Line35_stata2xcel.do"
	

***** ROUND 5 - OYO
	clear
	use "~/Dropbox (Gates Institute)/10_Nigeria/PMANG_Datasets/RoundOyo/Prelim100/NGROyo_WealthWeightAll_4May2018.dta", clear
	putexcel A36=("PMA2020-Oyo")
	putexcel B36=("Round 1")
	putexcel C36=("11/2017-05/2018")
	tab1 state strata ur

	** COUNT - Female Sample - All / Married Women  **
	preserve
	gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & last_night==1
	collapse (count) FQresponse_1
	mkmat FQresponse_1
	putexcel D36=matrix(FQresponse_1)
	restore
	preserve
	gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & last_night==1 & (FQmarital_status==1 | FQmarital_status==2)
	collapse (count) FQresponse_1
	mkmat FQresponse_1
	putexcel Q36=matrix(FQresponse_1)
	restore 

	*** Estimate Percentage and 95% CI
	keep if FRS_result==1 & HHQ_result==1 & last_night==1
	egen all=tag(FQmetainstanceID)
	egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)
	svyset Cluster_ID [pw=FQweight], singleunit(scaled)
	foreach group in all mar {
	preserve
		keep if `group'==1
		foreach indicator in cp mcp unmettot {
			svy: prop `indicator', citype(wilson)
			matrix reference=r(table)
			matrix `indicator'_`group'_percent=reference[1,2]*100	
			matrix `indicator'_`group'_se=reference[2,2]*100
			matrix `indicator'_`group'_ll=reference[5,2]*100
			matrix `indicator'_`group'_ul=reference[6,2]*100
			}	
		restore
		}
	run "/Users/asiewe/Documents/Data Analysis/DHS-PMA-Indicators/Descr_stats/Line36_stata2xcel.do"

	
/*
***********************************************************************************
*** 		UGANDA
***********************************************************************************
capture clear all
local excel "UG_KeyIndicators.xlsx"
local excel_sheet "CPR, mCPR, unmet need"
putexcel set "`excel'", modify sheet("`excel_sheet'")

***** ROUND 1
clear
use "~/Dropbox (Gates Institute)/11_Uganda/PMAUG_Datasets/Round1/Final_PublicRelease/HHQ/PMA2014_UGR1_HHQFQ_v1_29Dec2016/PMA2014_UGR1_HHQFQ_v1_29Dec2016.dta", clear
putexcel A8=("PMA2020")
putexcel B8=("Round 1")
putexcel C8=("5-6/2014")

** COUNT - Female Sample - All / Married Women  **
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel D8=matrix(FQresponse_1)
restore
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3) & (FQmarital_status==1 | FQmarital_status==2)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel Q8=matrix(FQresponse_1)
restore 

*** Estimate Percentage and 95% CI
keep if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
egen all=tag(FQmetainstanceID)
egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)

svyset EA_ID [pw=FQweight], strata(strata) singleunit(scaled)
foreach group in all mar {
preserve
	keep if `group'==1
	foreach indicator in cp mcp unmettot {
		svy: prop `indicator', citype(wilson)
		matrix reference=r(table)
		matrix `indicator'_`group'_percent=reference[1,2]*100	
		matrix `indicator'_`group'_se=reference[2,2]*100
		matrix `indicator'_`group'_ll=reference[5,2]*100
		matrix `indicator'_`group'_ul=reference[6,2]*100
		}	
	restore
	}
putexcel E8=matrix(cp_all_percent)
putexcel F8=matrix(cp_all_se)
putexcel G8=matrix(cp_all_ll)
putexcel H8=matrix(cp_all_ul)
putexcel I8=matrix(mcp_all_percent)
putexcel J8=matrix(mcp_all_se)
putexcel K8=matrix(mcp_all_ll)
putexcel L8=matrix(mcp_all_ul)
putexcel M8=matrix(unmettot_all_percent)
putexcel N8=matrix(unmettot_all_se)
putexcel O8=matrix(unmettot_all_ll)
putexcel P8=matrix(unmettot_all_ul)
putexcel R8=matrix(cp_mar_percent)
putexcel S8=matrix(cp_mar_se)
putexcel T8=matrix(cp_mar_ll)
putexcel U8=matrix(cp_mar_ul)
putexcel V8=matrix(mcp_mar_percent)
putexcel W8=matrix(mcp_mar_se)
putexcel X8=matrix(mcp_mar_ll)
putexcel Y8=matrix(mcp_mar_ul)
putexcel Z8=matrix(unmettot_mar_percent)
putexcel AA8=matrix(unmettot_mar_se)
putexcel AB8=matrix(unmettot_mar_ll)
putexcel AC8=matrix(unmettot_mar_ul)


***** ROUND 2
clear
use "~/Dropbox (Gates Institute)/11_Uganda/PMAUG_Datasets/Round2/Final_PublicRelease/HHQ/PMA2015_UGR2_HHQFQ_v1_29Dec2016/PMA2015_UGR2_HHQFQ_v1_29Dec2016.dta", clear
putexcel B9=("Round 2")
putexcel C9=("1-2/2015")

** COUNT - Female Sample - All / Married Women  **
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel D9=matrix(FQresponse_1)
restore
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3) & (FQmarital_status==1 | FQmarital_status==2)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel Q9=matrix(FQresponse_1)
restore 

** Estimate Percentage and 95% CI
keep if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
egen all=tag(FQmetainstanceID)
egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)

svyset EA_ID [pw=FQweight], strata(strata) singleunit(scaled)
foreach group in all mar {
preserve
	keep if `group'==1
	foreach indicator in cp mcp unmettot {
		svy: prop `indicator', citype(wilson)
		matrix reference=r(table)
		matrix `indicator'_`group'_percent=reference[1,2]*100	
		matrix `indicator'_`group'_se=reference[2,2]*100
		matrix `indicator'_`group'_ll=reference[5,2]*100
		matrix `indicator'_`group'_ul=reference[6,2]*100
		}	
	restore
	}
putexcel E9=matrix(cp_all_percent)
putexcel F9=matrix(cp_all_se)
putexcel G9=matrix(cp_all_ll)
putexcel H9=matrix(cp_all_ul)
putexcel I9=matrix(mcp_all_percent)
putexcel J9=matrix(mcp_all_se)
putexcel K9=matrix(mcp_all_ll)
putexcel L9=matrix(mcp_all_ul)
putexcel M9=matrix(unmettot_all_percent)
putexcel N9=matrix(unmettot_all_se)
putexcel O9=matrix(unmettot_all_ll)
putexcel P9=matrix(unmettot_all_ul)
putexcel R9=matrix(cp_mar_percent)
putexcel S9=matrix(cp_mar_se)
putexcel T9=matrix(cp_mar_ll)
putexcel U9=matrix(cp_mar_ul)
putexcel V9=matrix(mcp_mar_percent)
putexcel W9=matrix(mcp_mar_se)
putexcel X9=matrix(mcp_mar_ll)
putexcel Y9=matrix(mcp_mar_ul)
putexcel Z9=matrix(unmettot_mar_percent)
putexcel AA9=matrix(unmettot_mar_se)
putexcel AB9=matrix(unmettot_mar_ll)
putexcel AC9=matrix(unmettot_mar_ul)


***** ROUND 3
clear
use "~/Dropbox (Gates Institute)/11_Uganda/PMAUG_Datasets/Round3/Final_PublicRelease/HHQ/PMA2015_UGR3_HHQFQ_v1_29Dec2016/PMA2015_UGR3_HHQFQ_v1_29Dec2016.dta", clear
putexcel B10=("Round 3")
putexcel C10=("8-10/2015")
	
** COUNT - Female Sample - All / Married Women  **
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel D10=matrix(FQresponse_1)
restore
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3) & (FQmarital_status==1 | FQmarital_status==2)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel Q10=matrix(FQresponse_1)
restore 

** Estimate Percentage and 95% CI
keep if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
egen all=tag(FQmetainstanceID)
egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)

svyset EA_ID [pw=FQweight], strata(strata) singleunit(scaled)
foreach group in all mar {
preserve
	keep if `group'==1
	foreach indicator in cp mcp unmettot {
		svy: prop `indicator', citype(wilson)
		matrix reference=r(table)
		matrix `indicator'_`group'_percent=reference[1,2]*100	
		matrix `indicator'_`group'_se=reference[2,2]*100
		matrix `indicator'_`group'_ll=reference[5,2]*100
		matrix `indicator'_`group'_ul=reference[6,2]*100
		}	
	restore
	}
putexcel E10=matrix(cp_all_percent)
putexcel F10=matrix(cp_all_se)
putexcel G10=matrix(cp_all_ll)
putexcel H10=matrix(cp_all_ul)
putexcel I10=matrix(mcp_all_percent)
putexcel J10=matrix(mcp_all_se)
putexcel K10=matrix(mcp_all_ll)
putexcel L10=matrix(mcp_all_ul)
putexcel M10=matrix(unmettot_all_percent)
putexcel N10=matrix(unmettot_all_se)
putexcel O10=matrix(unmettot_all_ll)
putexcel P10=matrix(unmettot_all_ul)
putexcel R10=matrix(cp_mar_percent)
putexcel S10=matrix(cp_mar_se)
putexcel T10=matrix(cp_mar_ll)
putexcel U10=matrix(cp_mar_ul)
putexcel V10=matrix(mcp_mar_percent)
putexcel W10=matrix(mcp_mar_se)
putexcel X10=matrix(mcp_mar_ll)
putexcel Y10=matrix(mcp_mar_ul)
putexcel Z10=matrix(unmettot_mar_percent)
putexcel AA10=matrix(unmettot_mar_se)
putexcel AB10=matrix(unmettot_mar_ll)
putexcel AC10=matrix(unmettot_mar_ul)


**** ROUND 4
clear
use "~/Dropbox (Gates Institute)/11_Uganda/PMAUG_Datasets/Round4/Final_PublicRelease/HHQ/PMA2016_UGR4_HHQFQ_v1_29Dec2016/PMA2016_UGR4_HHQFQ_v1_29Dec2016.dta", clear
putexcel B11=("Round 4")
putexcel C11=("3-4/2016")

** COUNT - Female Sample - All / Married Women  **
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel D11=matrix(FQresponse_1)
restore
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3) & (FQmarital_status==1 | FQmarital_status==2)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel Q11=matrix(FQresponse_1)
restore 

** Estimate Percentage and 95% CI
keep if FRS_result==1 & HHQ_result==1 & (usual_member==1 | usual_member==3)
egen all=tag(FQmetainstanceID)
egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)

svyset EA_ID [pw=FQweight], strata(strata) singleunit(scaled)
foreach group in all mar {
preserve
	keep if `group'==1
	foreach indicator in cp mcp unmettot {
		svy: prop `indicator', citype(wilson)
		matrix reference=r(table)
		matrix `indicator'_`group'_percent=reference[1,2]*100	
		matrix `indicator'_`group'_se=reference[2,2]*100
		matrix `indicator'_`group'_ll=reference[5,2]*100
		matrix `indicator'_`group'_ul=reference[6,2]*100
		}	
	restore
	}
putexcel E11=matrix(cp_all_percent)
putexcel F11=matrix(cp_all_se)
putexcel G11=matrix(cp_all_ll)
putexcel H11=matrix(cp_all_ul)
putexcel I11=matrix(mcp_all_percent)
putexcel J11=matrix(mcp_all_se)
putexcel K11=matrix(mcp_all_ll)
putexcel L11=matrix(mcp_all_ul)
putexcel M11=matrix(unmettot_all_percent)
putexcel N11=matrix(unmettot_all_se)
putexcel O11=matrix(unmettot_all_ll)
putexcel P11=matrix(unmettot_all_ul)
putexcel R11=matrix(cp_mar_percent)
putexcel S11=matrix(cp_mar_se)
putexcel T11=matrix(cp_mar_ll)
putexcel U11=matrix(cp_mar_ul)
putexcel V11=matrix(mcp_mar_percent)
putexcel W11=matrix(mcp_mar_se)
putexcel X11=matrix(mcp_mar_ll)
putexcel Y11=matrix(mcp_mar_ul)
putexcel Z11=matrix(unmettot_mar_percent)
putexcel AA11=matrix(unmettot_mar_se)
putexcel AB11=matrix(unmettot_mar_ll)
putexcel AC11=matrix(unmettot_mar_ul)


**** ROUND 5
clear
use "~/Dropbox (Gates Institute)/11_Uganda/PMAUG_Datasets/Round5/Final_PublicRelease/HHQ/PMA2017_UGR5_HHQFQ_v1_8Feb2018/PMA2017_UGR5_HHQFQ_v1_8Feb2018.dta", clear
putexcel B12=("Round 5")
putexcel C12=("4-5/2017")

** COUNT - Female Sample - All / Married Women  **
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & last_night==1
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel D12=matrix(FQresponse_1)
restore
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & last_night==1 & (FQmarital_status==1 | FQmarital_status==2)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel Q12=matrix(FQresponse_1)
restore 

** Estimate Percentage and 95% CI
keep if FRS_result==1 & HHQ_result==1 & last_night==1
egen all=tag(FQmetainstanceID)
egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)

svyset EA_ID [pw=FQweight], strata(strata) singleunit(scaled)
foreach group in all mar {
preserve
	keep if `group'==1
	foreach indicator in cp mcp unmettot {
		svy: prop `indicator', citype(wilson)
		matrix reference=r(table)
		matrix `indicator'_`group'_percent=reference[1,2]*100	
		matrix `indicator'_`group'_se=reference[2,2]*100
		matrix `indicator'_`group'_ll=reference[5,2]*100
		matrix `indicator'_`group'_ul=reference[6,2]*100
		}	
	restore
	}
putexcel E12=matrix(cp_all_percent)
putexcel F12=matrix(cp_all_se)
putexcel G12=matrix(cp_all_ll)
putexcel H12=matrix(cp_all_ul)
putexcel I12=matrix(mcp_all_percent)
putexcel J12=matrix(mcp_all_se)
putexcel K12=matrix(mcp_all_ll)
putexcel L12=matrix(mcp_all_ul)
putexcel M12=matrix(unmettot_all_percent)
putexcel N12=matrix(unmettot_all_se)
putexcel O12=matrix(unmettot_all_ll)
putexcel P12=matrix(unmettot_all_ul)
putexcel R12=matrix(cp_mar_percent)
putexcel S12=matrix(cp_mar_se)
putexcel T12=matrix(cp_mar_ll)
putexcel U12=matrix(cp_mar_ul)
putexcel V12=matrix(mcp_mar_percent)
putexcel W12=matrix(mcp_mar_se)
putexcel X12=matrix(mcp_mar_ll)
putexcel Y12=matrix(mcp_mar_ul)
putexcel Z12=matrix(unmettot_mar_percent)
putexcel AA12=matrix(unmettot_mar_se)
putexcel AB12=matrix(unmettot_mar_ll)
putexcel AC12=matrix(unmettot_mar_ul)


**** ROUND 6
clear
use "~/Dropbox (Gates Institute)/11_Uganda/PMAUG_Datasets/Round6/Prelim100/UGR6_WealthWeightAll_6Jun2018.dta", clear
putexcel B13=("Round 6")
putexcel C13=("4-5/2018")

** COUNT - Female Sample - All / Married Women  **
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & last_night==1
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel D13=matrix(FQresponse_1)
restore
preserve
gen FQresponse_1=1 if FRS_result==1 & HHQ_result==1 & last_night==1 & (FQmarital_status==1 | FQmarital_status==2)
collapse (count) FQresponse_1
mkmat FQresponse_1
putexcel Q13=matrix(FQresponse_1)
restore 

** Estimate Percentage and 95% CI
keep if FRS_result==1 & HHQ_result==1 & last_night==1
egen all=tag(FQmetainstanceID)
egen mar=tag(FQmetainstanceID) if (FQmarital_status==1 | FQmarital_status==2)

svyset EA [pw=FQweight], strata(strata) singleunit(scaled)
foreach group in all mar {
preserve
	keep if `group'==1
	foreach indicator in cp mcp unmettot {
		svy: prop `indicator', citype(wilson)
		matrix reference=r(table)
		matrix `indicator'_`group'_percent=reference[1,2]*100	
		matrix `indicator'_`group'_se=reference[2,2]*100
		matrix `indicator'_`group'_ll=reference[5,2]*100
		matrix `indicator'_`group'_ul=reference[6,2]*100
		}	
	restore
	}
putexcel E13=matrix(cp_all_percent)
putexcel F13=matrix(cp_all_se)
putexcel G13=matrix(cp_all_ll)
putexcel H13=matrix(cp_all_ul)
putexcel I13=matrix(mcp_all_percent)
putexcel J13=matrix(mcp_all_se)
putexcel K13=matrix(mcp_all_ll)
putexcel L13=matrix(mcp_all_ul)
putexcel M13=matrix(unmettot_all_percent)
putexcel N13=matrix(unmettot_all_se)
putexcel O13=matrix(unmettot_all_ll)
putexcel P13=matrix(unmettot_all_ul)
putexcel R13=matrix(cp_mar_percent)
putexcel S13=matrix(cp_mar_se)
putexcel T13=matrix(cp_mar_ll)
putexcel U13=matrix(cp_mar_ul)
putexcel V13=matrix(mcp_mar_percent)
putexcel W13=matrix(mcp_mar_se)
putexcel X13=matrix(mcp_mar_ll)
putexcel Y13=matrix(mcp_mar_ul)
putexcel Z13=matrix(unmettot_mar_percent)
putexcel AA13=matrix(unmettot_mar_se)
putexcel AB13=matrix(unmettot_mar_ll)
putexcel AC13=matrix(unmettot_mar_ul)

