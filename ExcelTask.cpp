// -*-C++-*-
/*!
 * @file  ExcelTask.cpp
 * @brief
 *
 */


#include "stdafx.h"
#include "ExcelTask.h"
#include "myExcel.h"



ExcelTask::ExcelTask()
{
	
}


template <class T>
void getProperty(coil::Properties& prop, const char* key, T& value)
{
if (prop.findNode(key) != 0)
  {
    T tmp;
    if (coil::stringTo(tmp, prop[key].c_str()))
      {
        value = tmp;
      }
  }
}


int ExcelTask::svc()
{
	
	myExcel::Obj = gcnew myExcel();
	std::string filePath = "";
    coil::Properties& prop(::RTC::Manager::instance().getConfig());
    getProperty(prop, "excel.filename", filePath);
    myExcel::Obj->Open(gcnew System::String(filePath.c_str()));
	myExcel::Obj->filename = "";
	return 0;
}


