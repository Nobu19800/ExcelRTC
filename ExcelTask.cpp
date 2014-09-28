// -*-C++-*-
/*!
 * @file  ExcelTask.cpp
 * @brief Excelの起動、COMの参照をするスレッド
 *
 */


#include "stdafx.h"
#include "ExcelTask.h"
#include "myExcel.h"



ExcelTask::ExcelTask(ExcelControl *m_rtc)
{
	mrtc = m_rtc;
}





int ExcelTask::svc()
{
	
	
	myExcel::Obj = gcnew myExcel();
	std::string filePath = "";
    coil::Properties& prop(::RTC::Manager::instance().getConfig());
    getProperty(prop, "excel.filename", filePath);
	filePath = Replace(filePath, "/", "\\");
	
    myExcel::Obj->Open(gcnew System::String(filePath.c_str()));
	//myExcel::Obj->filename = "";

	

	mrtc->Load();
	return 0;
}


