library(rClr)
library(stringr)

CaseAlgo_x_2 <- function(LimitRemaining,prevPTD,injury_body_part,DevAge,Payroll,Curve_Mod){
	#the function name can be named anything here: it will be used in Excel formula via BERT addin
	#for example, "=R.CaseAlgo_x_2(arg1...)"
	

  f <- file.path('c:/git/DLL/', 'CaseAlgo_x_2.dll')  #the path to the DLL file
  f <- path.expand(f)
  stopifnot( file.exists(f) )
  clrLoadAssembly(f)
  obj<-clrNew("Model_CaseAlgo_x_2,CaseAlgo_x_2") 
	#two args:  first is Model_{fn name}
	#second is the name of the dll minus extension

#in this section, re-type all DotNET variables to matching types in R	
	#DotNET --> R
	#Double->as.numeric()
	#Integer->as.integer()
	#String->as.character()
	#Boolean-->as.logical()
	
  LimitRemaining<-as.numeric(LimitRemaining)
  prevPTD<-as.numeric(prevPTD)
  injury_body_part<-as.character(injury_body_part)
  #by default, strings in the DLL are padded at 30 characters
  injury_body_part<-str_pad(injury_body_part, 30, side = "right", pad = " ")
  DevAge<-as.numeric(DevAge)
  Payroll<-as.numeric(Payroll)


	
  #don't include a _default for the exposure variable, if any on this DLL
  prevPTD_default<-as.integer(0)
  injury_body_part_default<-as.integer(0)
  DevAge_default<-as.integer(0)
  Payroll_default<-as.integer(0)

  result<-clrCall(obj,'Predict',
					LimitRemaining,
					prevPTD,prevPTD_default,
					injury_body_part,injury_body_part_default,
					DevAge,DevAge_default,
					Payroll,Payroll_default,
					Curve_Mod)

  result

} 