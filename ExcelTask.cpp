// -*-C++-*-
/*!
 * @file  ExcelTask.cpp
 * @brief Excelの起動、COMの参照をするスレッド
 *
 */


#include "stdafx.h"
#include "ExcelTask.h"
#include "ExcelObject.h"


/**
*@brief Excelの起動、COMの参照をするスレッドのコンストラクタ
*/
ExcelTask::ExcelTask(ExcelControl *m_rtc)
{
	mrtc = m_rtc;
}




/**
*@brief スレッドを実行
* @return
*/
int ExcelTask::svc()
{
	
	
	ExcelObject::Obj = gcnew ExcelObject();
	std::string filePath = mrtc->getFileName();
	if(filePath == "NewFile")
	{
		filePath = "";
	}
    //coil::Properties& prop(::RTC::Manager::instance().getConfig());
    //getProperty(prop, "excel.filename", filePath);
	
	coil::replaceString(filePath,  "/", "\\");
	
    ExcelObject::Obj->Open(gcnew System::String(filePath.c_str()));
	//ExcelObject::Obj->filename = "";

	

	mrtc->load();
	return 0;
}


