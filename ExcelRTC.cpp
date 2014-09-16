// ExcelRTC.cpp : ���C�� �v���W�F�N�g �t�@�C���ł��B

#include "stdafx.h"
#include "Form1.h"
#include <rtm/Manager.h>
#include <iostream>
#include <string>
#include <stdlib.h>
#include "ExcelRTComponent.h"

using namespace ExcelRTC;



void MyModuleInit(RTC::Manager* manager)
{
  ExcelRTCInit(manager);
  RTC::RtcBase* comp;

  comp = manager->createComponent("ExcelControl");

  if (comp==NULL)
  {
	  
    std::cerr << "Component create failed." << std::endl;
    abort();
  }

  


  return;
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


[STAThreadAttribute]
int main(array<System::String ^> ^args)
{
	// �R���g���[�����쐬�����O�ɁAWindows XP �r�W���A�����ʂ�L���ɂ��܂�
	System::Windows::Forms::Application::EnableVisualStyles();
	System::Windows::Forms::Application::SetCompatibleTextRenderingDefault(false);

	RTC::Manager* manager;
	char* argv[] = {""};
	int argc = 0;
	manager = RTC::Manager::init(argc, argv);

	



	manager->init(argc, argv);
	manager->setModuleInitProc(MyModuleInit);
	manager->activateManager();
	manager->runManager(true);


	myExcel::Obj = gcnew myExcel();
	std::string filePath = "";
    coil::Properties& prop(::RTC::Manager::instance().getConfig());
    getProperty(prop, "excel.filename", filePath);
    myExcel::Obj->Open(gcnew System::String(filePath.c_str()));
	myExcel::Obj->filename = "";


	// ���C�� �E�B���h�E���쐬���āA���s���܂�
	System::Windows::Forms::Application::Run(gcnew Form1());
	return 0;
}
