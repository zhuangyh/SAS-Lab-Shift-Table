 /*====================================================================
| COMPANY           Bancova LLC
| PROJECT:          BANCOVA SAS TRAINING
| 
|
| PROGRAM:       	lab.yz.sas
| PROGRAMMER(S):    Yonghua Zhuang
| DATE:             07/27/2016
| PURPOSE:          Generate Lab Tables
|                        
| QC PROGRAMMER(S): 
| QC DATE:          
|
| INPUT FILE DIRECTORY(S): C:\bancova2016summer\data\raw      
| OUTPUT FILE DIRECTORY(S):C:\bancova2016summer\data\output       
| OUTPUT AND PRINT SPECIFICATIONS: Lab.RTF    
|
|
| REVISION HISTORY
| DATE     BY        COMMENTS
|
|
=====================================================================*/
 

*----------------------------------------------*;
* options, directories, librefs etc.
*----------------------------------------------*;

** clean the log and output screen; 
dm 'log; clear; output; clear';

** log, output and procedure options; 
options center formchar="|____|||___+=|_/\<>*" missing = '.' nobyline nodate;


** define the location for the input and output; 
libname raw "C:\bancova2016summer\data\raw" access=readonly;
libname derived "C:\bancova2016summer\data\derived";
%let outdir=C:\bancova2016summer\output;

*** create styles used for RTF-output ***;
proc template;
    DEFINE STYLE PANDA;
	PARENT= Styles.sasdocprinter;
	style fonts FROM FONTS /
		'titleFont' = ("courier new",12pt)
		'titleFont2'= ("courier new", 8pt)
		'headingFont' = ("times roman",10pt, bold)                            
        'docFont' = ("times roman",10pt);	
	style SystemFooter from systemfooter/
		font=fonts('titleFont2');
	replace body/
		bottommargin = 1in                                                
        topmargin = 1.5in                                                   
        rightmargin = 1in                                                 
        leftmargin = 1in; 	
	style TABLE from table/
		cellpadding=0
		cellspacing=0
		OUTPUTWIDTH=95%
		BORDERWIDTH=2PT;
	END;
run;

proc template;
    DEFINE STYLE lab1;
	PARENT= Styles.PANDA;
	style TABLE from table/
		cellpadding=0
		cellspacing=0
		rules=none
		frame=void;
	END;
run;

*** Create formats ***;
proc format;
	value treatmnt
	    1='Anticancer000'
		2='Anticancer001'
		3='Total';
	value test	
		1='WBC'
		2='RBC'
		3='Hgb' 
		4='Hct'
		5='Plateles'
		6='MCV'
		7="MCH" 
		8='Lymphoc.'
		9='Monos.' 
		10='RDW' 
		11='Basos';
	value visit
		0='Baseline'
		1='Visit 1'
		2='Visit 2'
		3='Visit 3'
		4='Visit 4'
		5='Visit 5'
		6='Visit 6'
		7='Visit 7'
		8='Visit 8'
		9='Visit 9'
		10='Visit 10'
		20='Visit 20';	
	value baseline
		1='Low'
		2='Normal'
		3='High'
		11='Low1'
		12='Normal1'
		13='High1';
run;


*****************************************;
*   Import Data and Data preparation    *;
*****************************************;
%macro import(xlsfile);
proc import out= work.&xlsfile 
            datafile= "C:\bancova2016summer\data\raw\&xlsfile..xls" 
            dbms=XLS replace;
run;

%mend import;
%import(hemadata); quit;
%import(demog_data); quit;

*** Select nonmising subjid and choose safety population ***; 
data dm;
	set demog_data;
	where ^missing(subjid);
	where safety=1;
run;

*** Delete observations wtihout any test results ***; 
data lab;
	set hemadata ;
	where ^missing(subjid);
	array num(*) BASOS HCT HGB LYMPHO MCH MCV MONOS PLAT RBC RDW WBC;   
	count=0;                          
   	do j = 1 to dim(num);           
      if not missing(num(j)) then count+1;     
  	end;   
   if count=0 then delete; 
   drop j count;   
run;

*** Find duplicates and keep unique observations ***; 
%macro dupfind (source=,var1=,var2=,lastvar=,dups=,uni=);
Proc sort data=&source;
		by &var1;
run;
data &uni &dups;
     set &source;
     by &var2;
     if not (first.&lastvar and last.&lastvar) then output &dups;
     if last.&lastvar then output &uni;
run;
%mend dupfind;
%dupfind(source=dm,var1=subjid, var2=subjid, lastvar=subjid, dups=demogdup, uni=dm1);
%dupfind(source=lab,var1=subjid visit, var2=subjid visit, lastvar=visit, dups=labdup, uni=lab1);

*** Merge dm and lab dataset ***;
data lab2 ;
   merge dm (keep=subjid treatmnt safety in=in1)  
		lab1 (drop=labdate in=in2);
   by subjid;
   if in1 and in2 then output lab2;
run ;

*** Transpose data ***;
proc transpose data=lab2 out=lab3(drop=_Name_ rename=(_Label_=test COL1=Value)) ;
	by subjid visit treatmnt;
	var BASOS HCT HGB LYMPHO MCH MCV MONOS PLAT RBC RDW WBC;
run;

*** Boxplot for result by tests and check outliers ***;
proc sgplot data=lab3;
   vbox value / category=test ;
   xaxis label="Test";
   keylegend / title="Boxplot for all test values";
run; 

*** Identify outliers ***;
proc univariate data=lab3 nextrval=5;
   var value;
   class test;
   where test in ('HCT' 'MCV' 'PLAT');
run;

*** Remove outliers ***;
Data lab4 outlier;
	set lab3;
	if (test='HCT' and value > 200) 
		or (test='MCV' and value <50) 
		then output outlier;
	else output lab4;
run;

*** Get baseline ***;
data baseline (rename=(value=value0));
	set lab4;
	if visit=0 then output baseline;
	drop visit;
run;

proc sort data=baseline;
	by treatmnt subjid test;
run;

proc sort data=lab4;
	by treatmnt subjid test;
run;

*** Caculate changes from baseline ***;
data lab4c;
	merge 	baseline (in=a)
			lab4 (in=b);
	by treatmnt subjid test;
	if visit > 0 then value=value-value0;
	if a and b and visit >0 then output lab4c;
run; 

*** add total treatment group ***;
data lab4;
	set lab4;
	output;
	treatmnt=3;
	output;
run;
data lab4c;
	set lab4c;
	output;
	treatmnt=3;
	output;
run;

data lab4c;
	set lab4c;
	output;
	treatmnt=3;
	output;
run;

*** Count the freq for each treatment group ***;
proc sql noprint;
	select count(distinct subjid) into :tt1-:tt3
	from lab4
	group by treatmnt
	order by treatmnt;
	quit;
%put &tt1, &tt2, &tt3;

proc sort data=lab4;
	by test subjid visit;
run;

proc sort data=lab4c;
	by test subjid visit;
run;

*** Creat a dataset for each visit header in tables ***;
data visit(keep=visit group id);
	set lab4;
	group=put(visit, visit.);
	id=visit;	
run;
%dupfind(source=visit,var1=visit, var2=visit, lastvar=visit, dups=visitdup, uni=visit1);

data visit2;
	set visit1;
	if visit = 0 then delete;
run;


*****************************************;
*   Generate Table 1&2    *;
*****************************************;
%macro labtable(sourcefile=, visith=, treatmnt=, treatmntg=, outtable=, title2=, title3=);

*** Get statistics for each test group by treatment ***;
proc summary data=&sourcefile; *lab4 lab4c;
	class visit test;
	var value; 
	where treatmnt=&treatmnt;
	output out=final1
		n		= c1
		mean	= c2
		median	= c3
		std		= c4
		min		= c5
		max		= c6;
run;

*** Delte  non-relevant results ***;
data final1 (drop=_TYPE_);
	set final1;
	where _TYPE_= 3;
run;

*** Cancatenate and format statistics as shown in table shell ***;
data final2 (drop=c1-c6 _freq_);
	set final1;
	stat1=strip(put(c1,3.));
	stat2=strip(put(c2,3.))||'('||strip(put(c3, 6.2))||')';
	stat3=strip(put(c4,3.));
	stat4=strip(put(c5,3.))||','||strip(put(c6, 3.));
run;

*** Transpose table to match table shell ***;
proc transpose data=final2 out=final3 ;
	by visit;
	var stat1-stat4;
	id test;
run;

*** Rename statistic names ***;
data final3;
	length group $ 20. id 8.;
	set final3;
	if _Name_='stat1' then do;
			group="  N";
			id=101;
			end;
	if _Name_='stat2' then do;
			group="  Mean(Std Dev)";
			id=102;
			end;
	if _Name_='stat3' then do;
			group="  Median";
			id=103;
			end;
	if _Name_='stat4' then do;
			group="  Minimum, Maximum";
			id=104;
			end;
run;

*** Add visit header for each visit group ***;
data final4 (drop=_Name_);
	length group $20;
	set 
		&visith
		final3;
run;

proc sort data=final4;
by visit id ;
run;

*** Reodrder column to match table shell ***;
data final4;
	retain group WBC RBC HGB HCT PLAT  MCV MCH LYMPHO MONOS RDW BASOS;
	set final4;
run;

*** Add a blank line for each visit group ***;
data final5 (drop=visit id) ; 
	do _n_=1 by 1 until(last.visit);
		set final4 end=last;
		by visit notsorted;
		output;
	end;
	call missing(of _all_);
	if not last then output;
run;

***** write table 1&2 to RTF using proc report*****;
ods listing close;
options nodate nonumber orientation=landscape missing='';
ods escapechar='^';
ods rtf file = "&outdir\&outtable" style=lab1;
proc report data=final5 nowindows headline missing headskip split='|'
	style(header)=[background=white borderbottomcolor=black borderbottomwidth=2];
	column	((" " group)
			(" ^S={borderbottomcolor=black borderbottomwidth=2} &treatmntg (N=&&&tt&treatmnt.)" 
				WBC RBC HGB HCT PLAT MCV LYMPHO MONOS RDW BASOS ));
		
	define 	group/display 			style(header)=[just=l vjust=top asis=on bordertopcolor=white]
									style(column)={just=l vjust=bottom asis=on cellwidth=15%}
									"  |Visit";
	define WBC /display 			style(header)=[just=c vjust=top]
									style(column)=[just=c vjust=bottom  cellwidth=8%]
									"WBC | x10^{super 6}/L";
	define RBC/display 				style(header)=[just=c vjust=top]
									style(column)=[just=c vjust=bottom  cellwidth=8%]
									"RBC | x10^{super 9}/L";
	define HGB/display 			style(header)=[just=c vjust=top ]
									style(column)=[just=c vjust=bottom cellwidth=8%]
									"Hgb | g/dL" ;
	define HCT/display 			style(header)=[just=c vjust=top ]
									style(column)=[just=c vjust=bottom  cellwidth=8%]
									"Hct | Ratio";
	define PLAT/display 		style(header)=[just=c vjust=top ]
									style(column)=[just=c vjust=bottom cellwidth=8%]
									"Platelets | x10^{super 6}/L";
	define MCV/display 			style(header)=[just=c vjust=top ]
									style(column)=[just=c vjust=bottom  cellwidth=8%]
									"MCV | U/L";
	define LYMPHO /display 			style(header)=[just=c vjust=top]
									style(column)=[just=c vjust=bottom cellwidth=8%]
									"Lymphoc. | U/L";

	define MONOS /display 			style(header)=[just=c vjust=top]
									style(column)=[just=c vjust=bottom cellwidth=8%]
									"Monos. | U/L";

	define RDW /display 			style(header)=[just=c vjust=top ]
									style(column)=[just=c vjust=bottom  cellwidth=8%]
									"RDW | U/L";
	define BASOS/display 			style(header)=[just=c vjust=top ]
									style(column)=[just=c vjust=bottom  cellwidth=8%]
									"Basos. | U/L";
	
	title1  h=8pt j=l "Bancova_lab" j=r "page^{pageof}";
	title2 h=10pt j=l  &title2;
  	title3  h=10pt  j=l &title3;
  	title4  h=10pt "(SAFETY Population)";
	footnote1 "^R'\brdrb\brdrs\brdrw1";
	footnote2 "lab_yz.sas  submitted  &sysdate9. at  &systime by Yonghua Zhuang";
run ; 

%let text= %str(^S={just=l font=('courier new',8pt)}); 
	ods rtf text=' ';
	
ods rtf close;
ods listing;
%MEND labtable;

%labtable(	sourcefile=lab4, visith=visit1, treatmnt=1, treatmntg=Anticancer00, outtable=Table1_treat1.rtf, 
			title2="Table 6.1", 
			title3= "SUMMARY OF Laboratory Parameters Over Time: Hematology" );

%labtable(	sourcefile=lab4, visith=visit1, treatmnt=2, treatmntg=Anticancer01, outtable=Table1_treat2.rtf, 
			title2="Table 6.1", 
			title3= "SUMMARY OF Laboratory Parameters Over Time: Hematology" );

%labtable(	sourcefile=lab4,visith=visit1, treatmnt=3, treatmntg=Total, outtable=Table1_total.rtf, 
			title2="Table 6.1", 
			title3= "SUMMARY OF Laboratory Parameters Over Time: Hematology" );

%labtable(	sourcefile=lab4c, visith=visit2, treatmnt=1, treatmntg=Anticancer00, outtable=Table2_treat1.rtf, 
			title2="Table 6.2", 
			title3= "SUMMARY OF Changes from Baseline Clinical Laboratory Parameters Over Time: Hematology" );

%labtable(	sourcefile=lab4c, visith=visit2, treatmnt=2, treatmntg=Anticancer01, outtable=Table2_treat2.rtf, 
			title2="Table 6.2", 
			title3= "SUMMARY OF Changes from Baseline Clinical Laboratory Parameters Over Time: Hematology" );


*************************************************************
*          Make Table 3 (lift table)                        *       
*************************************************************;

*****Create dummy data frame to have a whole structure*****;
data dummy;
	set visit1(keep=visit);
		by visit;
			do treatmnt=1 to 2;
				do baseline=1 to 3;
					do range=1 to 3;
					output;
					end;
				end;
			end;
run;

%macro shifttable(test=, outtable=);
*****Choose selected test for later analyis *****;
*****Delete total treatment grouo *****;
data &test;
	set lab4;
	where test="&test" and treatmnt ^=3;
run;

*****Get mean*****;
proc sql;
  select mean(value) into:st1
    from wbc;
QUIT;

*****Get STD*****;
proc sql;
  select std(value) into:st2
    from wbc;
QUIT;

*****Define threshold and classify to 1-3 (low, normal, high)*****;
data &test;
	set &test;
	upthreshold 	= &st1 + 1.5 * &st2; *Instead "mean+3*STD" for better display in final tables; 
	lowthreshold 	= &st1 - 1.5 * &st2;
	if value < lowthreshold then range = 1;
		else if value > upthreshold then range = 2;
		else if ^missing(value) then range =3;
run;

*****Get baseline for each subject*****;
data baseline1(rename=(range=baseline)drop=visit);
	set &test;
	if visit=0 then output baseline1;
run;

proc sort data=baseline1;
	by treatmnt subjid;
run;

proc sort data=&test;
	by treatmnt subjid visit;
run;

*****Add baseline for each subject*****;
data temp1 (drop=Value upthreshold lowthreshold) ;
	merge 	baseline1 (in=a)
			&test (in=b);
	by treatmnt subjid;
	if a and b and visit > 0 then output;
run; 

proc sort data=temp1;
	by visit;
run;

*****Count N and get percent for each group by treatment and visit*****;
proc tabulate data=temp1 out=temp2;
  	class  baseline range ;
  	table  
		baseline='baseline',
		range=''*(n='n'  PCTN='p')*F=10./ RTS=13.;
	by visit treatmnt;
run;

proc sort data=temp2;
	by visit treatmnt;
run;

*****Contecate with dummy dataframe*****;
data temp3;
	set 	dummy
			temp2;
run;

proc sort data=temp3 (drop=_TYPE_ _PAGE_ _TABLE_ );
	by visit treatmnt baseline range;
run;

*****If not missing, delete corresponding rows from dummy data frame*****;
data temp4;
	set temp3;
	by visit treatmnt baseline range;
	if first.range=1 and last.range=0 then delete;
run;

*****Fill 0 to all missing values*****;
data temp5;
	set temp4;
	array change _numeric_;
        do over change;
            if change=. then change=0;
        end;
 run;

 *****Differentiate baseline variable for two treatment groups*****;
data temp6;
	set temp5;
	if treatmnt = 2 then baseline=baseline+10;
run;

*****Transpose value, baseline grouped by visit*****;
data temp7 (drop=treatmnt N PctN_00);
	set temp6;
	value=strip(put(N,3.))||'('||strip(put(PctN_00, 5.1))||'%)'; 
	format baseline baseline.;
run;

proc sort data=temp7;
	by visit range;
run;

*****Transpose value, baseline grouped by visit*****;
proc transpose data=temp7 out=temp8 (drop=_NAME_);
	by visit range;
	var value;
	id baseline;
run;

*****Add a visit header line for each visit group*****;
data temp9;
	set 
		visit1(keep=visit group)
		temp8;
run;

proc sort data=temp9;
	by visit range;
run;

*****Recode range to a character variable*****;
data temp10 (drop=range);
	length group $10.;
	retain group Low Normal High Low1 Normal1 High1;
	set temp9;
	if range=1 then group="  Low";
	if range=2 then group="  Normal";
	if range=3 then group="  High";
	where visit >0;
run;

*****Add a blank line between each visit group*****;
data temp11 (drop=visit) ; 
	do _n_=1 by 1 until(last.visit);
		set temp10 end=last;
		by visit notsorted;
		output;
	end;
	call missing(of _all_);
	if not last then output;
run;

***** write table 3 to RTF using proc report*****;
ods listing close;
options nodate nonumber orientation=landscape missing='';
ods rtf file = "&outdir\&outtable" style=lab1;
proc report data=temp11 nowindows headline missing headskip split='|'
	style(header)=[background=white borderbottomcolor=black borderbottomwidth=2];
	column	(
			('^S={borderbottomcolor=white}'
				('^S={borderbottomcolor=white}'
				group))
			("^S={just=l} Lab Test &test"
				(" Anticancer00  "
				Low Normal High)
				(" Anticancer01  "
				Low1 Normal1 High1)));
		
	define 	group/display 			style(header)=[just=l vjust=top asis=on  bordertopcolor=white]
									style(column)={just=l vjust=bottom asis=on  cellwidth=15%}
									"  Baseline";
	define Low /display 			style(header)=[just=c vjust=top bordertopcolor=white]
									style(column)=[just=c vjust=bottom  cellwidth=8%]
									"Low";
	define Normal/display 			style(header)=[just=c vjust=top bordertopcolor=white]
									style(column)=[just=c vjust=bottom  cellwidth=8%]
									"Normal";
	define High/display 			style(header)=[just=c vjust=top bordertopcolor=white]
									style(column)=[just=c vjust=bottom cellwidth=8%]
									"High";
	
	define Low1/display 			style(header)=[just=c vjust=top bordertopcolor=white]
									style(column)=[just=c vjust=bottom  cellwidth=8%]
									"Low";
	define Normal1/display 				style(header)=[just=c vjust=top bordertopcolor=white]
									style(column)=[just=c vjust=bottom  cellwidth=8%]
									"Normal";
	define High1/display 			style(header)=[just=c vjust=top bordertopcolor=white]
									style(column)=[just=c vjust=bottom cellwidth=8%]
									"High";

	title1  h=8pt j=l "Bancova_lab" j=r "page^{pageof}";
	title2 j=l "Table 6.3";
  	title3 j=l "Shift Table for Clinical Laboratory Parameters Over Time: Hematology";
  	title4 "(SAFETY Population)";
	footnote1 "^R'\brdrb\brdrs\brdrw1";
	footnote2 "lab_yz.sas  submitted  &sysdate9. at  &systime by Yonghua Zhuang";
run ; 

%let text= %str(^S={just=l font=('courier new',8pt)}); 
	ods rtf text=' ';
	
ods rtf close;
ods listing;
%MEND shifttable;

*****Generate Table 3 for each test*****;
%shifttable(test=WBC, outtable=Table3_WBC.rtf);
%shifttable(test=RBC, outtable=Table3_RBC.rtf);
/*%shifttable(test=HGB, outtable=Table3_HGB.rtf);*/
/*%shifttable(test=HCT, outtable=Table3_HCT.rtf);*/
/*%shifttable(test=PLAT, outtable=Table3_PLAT.rtf);*/
/*%shifttable(test=MCV, outtable=Table3_MCV.rtf);*/
/*%shifttable(test=MCH, outtable=Table3_MCH.rtf);*/
/*%shifttable(test=LYMPHO, outtable=Table3_LYMPHO.rtf);*/
/*%shifttable(test=MONOS, outtable=Table3_MONOS.rtf);*/
/*%shifttable(test=RDW, outtable=Table3_RDW.rtf);*/
/*%shifttable(test=BASOS, outtable=Table3_BASOS.rtf);*/
