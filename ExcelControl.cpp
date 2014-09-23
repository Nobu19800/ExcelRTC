// -*-C++-*-
/*!
 * @file  ExcelControl.cpp
 * @brief
 *
 */

#include "stdafx.h"
#include <rtm/Manager.h>
#include <iostream>
#include <string>
#include <stdlib.h>
#include "ExcelControl.h"


#ifdef CREATE_EXE_EXCELRTC
	using namespace ExcelRTC;
#endif





#include "stdafx.h"
#include "ExcelControl.h"

#include <iostream>
#include <fstream>
#include "MyObject.h"



 
using namespace std;


ExcelControl *tertc;

static const char* excelrtc_spec[] =
  {
    "implementation_id", "ExcelControl",
    "type_name",         "ExcelControl",
    "description",       "Excel Control Component",
    "version",           "0.0.2",
    "vendor",            "Miyamoto Nobuhiko",
    "category",          "Category",
    "activity_type",     "PERIODIC",
    "kind",              "DataFlowComponent",
    "max_instance",      "1",
    "language",          "C++",
    "lang_type",         "compile",
	"conf.default.file_path", "NewFile",
	"conf.default.actionLock", "0",
    "conf.default.Red", "255",
    "conf.default.Green", "255",
    "conf.default.Blue", "0",
	"conf.dataport0.port_type", "DataInPort",
	"conf.dataport0.data_type", "TimedFloat",
	"conf.dataport0.column", "1",
	"conf.dataport0.start_row", "A",
	"conf.dataport0.end_row", "A",
	"conf.dataport0.sheetname", "Sheet1",
	"conf.__widget__.actionLock", "radio",
    "conf.__widget__.Red", "spin",
    "conf.__widget__.Green", "spin",
    "conf.__widget__.Blue", "spin",
	"conf.__widget__.file_path", "text",
	"conf.__widget__.port_type", "radio",
	"conf.__widget__.column", "spin",
	"conf.__widget__.start_row", "text",
	"conf.__widget__.end_row", "text",
	"conf.__widget__.sheetname", "text",
	"conf.__widget__.data_type", "radio",
	"conf.__constraints__.actionLock", "(0,1)",
    "conf.__constraints__.Red", "0<=x<=255",
    "conf.__constraints__.Green", "0<=x<=255",
    "conf.__constraints__.Blue", "0<=x<=255",
    "conf.__constraints__.column", "1<=x<=1000",
	"conf.__constraints__.column", "1<=x<=1000",
	"conf.__constraints__.port_type", "(DataInPort,DataOutPort)",
	"conf.__constraints__.data_type", "(TimedDouble,TimedLong,TimedFloat,TimedShort,TimedULong,TimedUShort,TimedChar,TimedWChar,TimedBoolean,TimedOctet,TimedString,TimedWString,TimedDoubleSeq,TimedLongSeq,TimedFloatSeq,TimedShortSeq,TimedULongSeq,TimedUShortSeq,TimedCharSeq,TimedWCharSeq,TimedOctetSeq,TimedStringSeq,TimedWStringSeq)",
	""
  };







void SetTree(TreeObject *to)
{
	
	
	//System::String^ a = gcnew System::String(to->name.c_str());
	//System::Windows::Forms::MessageBox::Show(a);

	

	for(int i=0;i < to->to.size();i++)
	{
		SetTree(to->to[i]);
	}
}





ExcelControl::ExcelControl(RTC::Manager* manager)
  : RTC::DataFlowComponentBase(manager),
  m_SpreadSheetPort("SpreadSheet")

{
	tertc = this;
	m_manager = manager;
	/*string IP_adress="localhost:2809";
	RTC::CorbaNaming namingserver(manager->getORB(), IP_adress.c_str());

	vector<OtherPort> rtclist;
	TreeObject *to = new TreeObject(IP_adress);
    rtc_get_rtclist(namingserver,rtclist,to);*/
	//SetTree(to);

	/*for(int i=0;i < rtclist.size();i++)
	{
		CreatePort(rtclist[i]);
	}*/

	et = new ExcelTask(this);
	et->activate();
	

	
}

ExcelControl::~ExcelControl()
{
}

TreeObject* ExcelControl::GetRTCTree(string IP_adress)
{
	try
	{
		RTC::CorbaNaming namingserver(m_manager->getORB(), IP_adress.c_str());
		rtclist.clear();
		DeleteAllPort();
		TreeObject *to = new TreeObject(IP_adress);
		rtc_get_rtclist(namingserver,rtclist,to,IP_adress);

		return to;
	}
	catch(...)
	{
#ifdef CREATE_EXE_EXCELRTC
		System::Windows::Forms::MessageBox::Show("ネーミングサービスへの接続に失敗しました");
#endif
		return NULL;
	}
	
	//SetTree(to);
	
	
	
	
	
}

RTC::ReturnCode_t ExcelControl::onInitialize()
{
	//myExcel::Obj = gcnew myExcel();
  
  this->addConfigurationSetListener(ON_SET_CONFIG_SET, new MyConfigUpdateParam(this));


  bindParameter("file_path", file_path, "NewFile");
  bindParameter("data_type", conf_data_type, "TimedFloat");
  bindParameter("port_type", conf_port_type, "DataInPort");
  bindParameter("column", conf_column, "1");
  bindParameter("start_row", conf_start_row, "A");
  bindParameter("end_row", conf_end_row, "A");
  bindParameter("sheetname", conf_sheetname, "Sheet1");
  bindParameter("actionLock", actionLock, "0");
  bindParameter("Red", Red, "255");
  bindParameter("Green", Green, "255");
  bindParameter("Blue", Blue, "0");
  

  
  m_SpreadSheetPort.registerProvider("spreadsheet", "SpreadSheet::mSpreadSheet", m_spreadsheet);

  addPort(m_SpreadSheetPort);

  

 
  

  return RTC::RTC_OK;
}


RTC::ReturnCode_t ExcelControl::onDeactivated(RTC::UniqueId ec_id)
{
	ResetAllPort();

	return RTC::RTC_OK;
}


RTC::ReturnCode_t ExcelControl::onFinalize()
{
  myExcel::Obj->Close();
  return RTC::RTC_OK;
}

void ExcelControl::update_cellName()
{
	for(int i=0;i < ConfInPorts.size();i++)
	{
		ConfInPorts[i]->update_cellName();
		
	}

	for(int i=0;i < ConfOutPorts.size();i++)
	{
		ConfOutPorts[i]->update_cellName();

	}


	for(int i=0;i < InPorts.size();i++)
	{
		InPorts[i]->update_cellName();
		
		
	}

	for(int i=0;i < OutPorts.size();i++)
	{
		
		OutPorts[i]->update_cellName();

	}
}




RTC::ReturnCode_t ExcelControl::onExecute(RTC::UniqueId ec_id)
{
	if(actionLock == 1)
	{
		myExcel::Obj->xlApplication->ScreenUpdating = false;
		_mutex.lock();
	}

	myExcel::Obj->SetColor(Red, Green, Blue);

	for(int i=0;i < ConfInPorts.size();i++)
	{
		ConfInPorts[i]->PutData(true);
		
	}

	for(int i=0;i < ConfOutPorts.size();i++)
	{
		ConfOutPorts[i]->PutData(true);

	}


	for(int i=0;i < InPorts.size();i++)
	{
		if(InPorts[i]->attachPort.size() == 0)
		{
			InPorts[i]->PutData(true);
			
		}
		
		
	}

	for(int i=0;i < OutPorts.size();i++)
	{
		if(OutPorts[i]->attachPort.size() == 0)
		{
			OutPorts[i]->PutData(true);
			
		}
	}

	if(actionLock == 1)
	{
		myExcel::Obj->xlApplication->ScreenUpdating = true;
		_mutex.unlock();
	}
	

	
	
	

  return RTC::RTC_OK;
}

void ExcelControl::updateAPort(MyPortBase* ip)
{
	for(int i=0;i < ip->attachPort.size();i++)
	{
		MyPortBase *op = GetOutPort(ip->attachPort[i]);
		if(op)
		{
			if(op->attachPort.size() > 0)
			{
				bool in = true;
				for(int j=0;j < op->attachPort.size();j++)
				{
					MyPortBase *mpb = GetInPort(op->attachPort[j]);
					if(mpb != NULL)
					{
						if(mpb->isNew() == false)
						{
							in = false;
						}
					}
				}

				if(in)
				{
					_mutex.lock();
	
	
					for(int j=0;j < op->attachPort.size();j++)
					{
						MyPortBase *mpb = GetInPort(op->attachPort[j]);
						if(mpb != NULL)
						{
							mpb->PutData(false);
						}
					}
					op->PutData(false);

					_mutex.unlock();
				}
			
			

			}
		}
	}
}


void ExcelControl::DelDPort(std::vector<std::string> pt)
{
	for(int i=0;i < rtclist.size();i++)
	{
		if(rtclist[i].buff == pt && rtclist[i].mpb != NULL)
		{
			DeleteOtherPort(rtclist[i]);
		}

	}
}

void ExcelControl::ConfigUpdate()
{

	this->m_configsets.update();
	std::string sfn = Replace(file_path, "/", "\\");
	System::String ^tfn = gcnew System::String(sfn.c_str());
	//System::Console::WriteLine(tfn);
	
	if(sfn ==  "NewFile")
	{
		myExcel::Obj->Open("");
		Load();
	}
	else if(myExcel::Obj->filename != tfn)
	{
		
		myExcel::Obj->Open(tfn);
		Load();
		 
		
	}
	for(int i=0;i < 100;i++)
	{
		std::string dn = "dataport";
		std::stringstream ss;
		ss << i+1;
		dn = dn + ss.str();

		
		
		if(this->m_configsets.haveConfig(dn.c_str()))
		{
			this->m_configsets.activateConfigurationSet(dn.c_str());
			this->m_configsets.update();

			std::string tdt = "DataInPort";

			if(conf_column > 0 && conf_start_row.size() > 0)
			{

				MyPortBase* tmp = GetConfInPort(dn);
				if(tmp == NULL)
				{
					tmp = GetConfOutPort(dn);
					tdt = "DataOutPort";
				}
				if(tmp && tmp->data_type == conf_data_type && conf_port_type == tdt)
				{
					tmp->SetExcelParam(conf_column, conf_start_row, conf_sheetname, conf_end_row, true);
				}
				else
				{
					if(tmp)
					{
						tmp->pb->disconnect_all();
						removePort(*tmp->pb);
						
						std::vector<MyPortBase*>::iterator end_it = remove( ConfInPorts.begin(), ConfInPorts.end(), tmp );
						ConfInPorts.erase( end_it, ConfInPorts.end() );

						std::vector<MyPortBase*>::iterator end_it2 = remove( ConfOutPorts.begin(), ConfOutPorts.end(), tmp );
						ConfOutPorts.erase( end_it2, ConfOutPorts.end() );
					}
					
					
					if(conf_data_type == "TimedDouble")
					{
						
						ConfcrPort<RTC::TimedDouble, double>(dn, conf_port_type, conf_data_type, conf_column, conf_start_row, conf_sheetname, conf_end_row, true);
					}
					else if(conf_data_type == "TimedLong")
					{
						ConfcrPort<RTC::TimedLong, long>(dn, conf_port_type, conf_data_type, conf_column, conf_start_row, conf_sheetname, conf_end_row, true);
					}
					

					else if(conf_data_type == "TimedFloat")
					{
						
						ConfcrPort<RTC::TimedFloat, float>(dn, conf_port_type, conf_data_type, conf_column, conf_start_row, conf_sheetname, conf_end_row, true);
					}
					else if(conf_data_type == "TimedShort")
					{
						
						ConfcrPort<RTC::TimedShort, short>(dn, conf_port_type, conf_data_type, conf_column, conf_start_row, conf_sheetname, conf_end_row, true);
					}
					else if(conf_data_type == "TimedULong")
					{
						ConfcrPort<RTC::TimedULong, long>(dn, conf_port_type, conf_data_type, conf_column, conf_start_row, conf_sheetname, conf_end_row, true);
					}
					
					else if(conf_data_type == "TimedUShort")
					{
						ConfcrPort<RTC::TimedUShort, short>(dn, conf_port_type, conf_data_type, conf_column, conf_start_row, conf_sheetname, conf_end_row, true);
					}
					else if(conf_data_type == "TimedChar")
					{
						ConfcrPort<RTC::TimedChar, char>(dn, conf_port_type, conf_data_type, conf_column, conf_start_row, conf_sheetname, conf_end_row, true);
					}
					else if(conf_data_type == "TimedWChar")
					{
						//ConfcrPort<RTC::TimedWChar, char>(dn, conf_port_type, conf_data_type, conf_column, conf_start_row, conf_sheetname, conf_end_row, true);
					}
					else if(conf_data_type == "TimedBoolean")
					{
						ConfcrPort<RTC::TimedBoolean, bool>(dn, conf_port_type, conf_data_type, conf_column, conf_start_row, conf_sheetname, conf_end_row, true);
					}
					else if(conf_data_type == "TimedOctet")
					{
						
						ConfcrPort<RTC::TimedOctet, char>(dn, conf_port_type, conf_data_type, conf_column, conf_start_row, conf_sheetname, conf_end_row, true);
					}
					else if(conf_data_type == "TimedString")
					{
						
						//ConfcrPort<RTC::TimedString, char*>(dn, conf_port_type, conf_data_type, conf_column, conf_start_row, conf_sheetname, conf_end_row, true);
					}
					else if(conf_data_type == "TimedWString")
					{
						
						//ConfcrPort<RTC::TimedWString, char*>(dn, conf_port_type, conf_data_type, conf_column, conf_start_row, conf_sheetname, conf_end_row, true);
					}
					

					else if(conf_data_type == "TimedDoubleSeq")
					{
						
						ConfcrPortSeq<RTC::TimedDoubleSeq, double>(dn, conf_port_type, conf_data_type, conf_column, conf_start_row, conf_sheetname, conf_end_row, true);
					}
					else if(conf_data_type == "TimedLongSeq")
					{
						ConfcrPortSeq<RTC::TimedLongSeq, long>(dn, conf_port_type, conf_data_type, conf_column, conf_start_row, conf_sheetname, conf_end_row, true);
					}

					else if(conf_data_type == "TimedFloatSeq")
					{
						ConfcrPortSeq<RTC::TimedFloatSeq, float>(dn, conf_port_type, conf_data_type, conf_column, conf_start_row, conf_sheetname, conf_end_row, true);
					}
					
					else if(conf_data_type == "TimedShortSeq")
					{
						ConfcrPortSeq<RTC::TimedShortSeq, short>(dn, conf_port_type, conf_data_type, conf_column, conf_start_row, conf_sheetname, conf_end_row, true);
					}
					else if(conf_data_type == "TimedULongSeq")
					{
						
						ConfcrPortSeq<RTC::TimedULongSeq, long>(dn, conf_port_type, conf_data_type, conf_column, conf_start_row, conf_sheetname, conf_end_row, true);
					}
					else if(conf_data_type == "TimedUShortSeq")
					{
						ConfcrPortSeq<RTC::TimedUShortSeq, short>(dn, conf_port_type, conf_data_type, conf_column, conf_start_row, conf_sheetname, conf_end_row, true);
					}
					else if(conf_data_type == "TimedCharSeq")
					{
						ConfcrPortSeq<RTC::TimedCharSeq, char>(dn, conf_port_type, conf_data_type, conf_column, conf_start_row, conf_sheetname, conf_end_row, true);
					}
					else if(conf_data_type == "TimedWCharSeq")
					{
						//ConfcrPortSeq<RTC::TimedWCharSeq, char>(dn, conf_port_type, conf_data_type, conf_column, conf_start_row, conf_sheetname, conf_end_row, true);
					}
					
					else if(conf_data_type == "TimedOctetSeq")
					{
						ConfcrPortSeq<RTC::TimedOctetSeq, char>(dn, conf_port_type, conf_data_type, conf_column, conf_start_row, conf_sheetname, conf_end_row, true);
					}

					else if(conf_data_type == "TimedStringSeq")
					{
						//ConfcrPortSeq<RTC::TimedStringSeq, char*>(dn, conf_port_type, conf_data_type, conf_column, conf_start_row, conf_sheetname, conf_end_row, true);
					}

					else if(conf_data_type == "TimedWStringSeq")
					{
						//ConfcrPortSeq<RTC::TimedWStringSeq, char*>(dn, conf_port_type, conf_data_type, conf_column, conf_start_row, conf_sheetname, conf_end_row, true);
					}
				}	
					
			}
		}
	}
	
}

MyPortBase* ExcelControl::GetDPort(std::vector<std::string> pt)
{
	for(int i=0;i < rtclist.size();i++)
	{
		if(rtclist[i].buff == pt && rtclist[i].mpb != NULL)
		{
			return rtclist[i].mpb;
		}

	}
	return NULL;
}

MyPortBase* ExcelControl::SetDPort(std::vector<std::string> pt, int c, std::string l, std::string sn, std::string leng, bool mstate, bool msflag)
{
	/*for(int i=0;i < pt.size();i++)
	{
		std::cout << pt[i] << std::endl;
	}*/
	
	for(int i=0;i < rtclist.size();i++)
	{
		
		/*for(int j=0;j < rtclist[i].buff.size();j++)
		{
			std::cout << rtclist[i].buff[j] << std::endl;
			
		}
		std::cout<< std::endl;*/

		if(rtclist[i].buff == pt && rtclist[i].mpb == NULL)
		{
			System::String^ tmp = gcnew System::String(rtclist[i].buff[rtclist[i].buff.size()-2].c_str());
			tmp += "の";
			tmp += gcnew System::String(rtclist[i].buff[rtclist[i].buff.size()-1].c_str());
			tmp += "と通信するデータポートを作成しました。";

			MyPortBase* m_pb = CreatePort(rtclist[i], c, l, sn, leng, mstate);
			if(msflag)
			{
#ifdef CREATE_EXE_EXCELRTC
				System::Windows::Forms::MessageBox::Show(tmp);
#endif
				m_pb->update_cellName();
				
			}


			return m_pb;
		}
		else if(rtclist[i].buff == pt && rtclist[i].mpb != NULL)
		{
			rtclist[i].mpb->SetExcelParam(c, l, sn, leng, mstate);

			MyPortBase* m_pb = rtclist[i].mpb;
			if(msflag)
			{
				m_pb->update_cellName();
				
			}
			return m_pb;
		}
		
		/*for(int j=0;j < pt.size();j++)
		{
			std::cout << pt[j] << std::endl;
		}
		std::cout << std::endl;
		for(int j=0;j < rtclist[i].buff.size();j++)
		{
			std::cout << rtclist[i].buff[j] << std::endl;
		}
		std::cout << std::endl;*/

	}
	return NULL;
}

void ExcelControl::DeleteOtherPort(OtherPort &op)
{
	

	op.mpb->pb->disconnect_all();
	removePort(*op.mpb->pb);

	
	std::vector<MyPortBase*>::iterator end_it = remove( InPorts.begin(), InPorts.end(), op.mpb );
	InPorts.erase( end_it, InPorts.end() );

	std::vector<MyPortBase*>::iterator end_it2 = remove( OutPorts.begin(), OutPorts.end(), op.mpb );
	OutPorts.erase( end_it2, OutPorts.end() );

	delete op.mpb;
	op.mpb = NULL;

	Save();
#ifdef CREATE_EXE_EXCELRTC
	System::Windows::Forms::MessageBox::Show("削除しました");
#endif
	
}

void ExcelControl::AttachPort(MyPortBase *mpb, std::string n)
{
	
	MyPortBase *ip = GetInPort(n);

	if(ip != NULL && GetInPort(mpb->name) == NULL)
	{
		
		ip->attachPort.push_back(mpb->name);
		mpb->attachPort.push_back(n);

		System::String^ tmp = gcnew System::String(n.c_str());
		tmp += "と";
		tmp += gcnew System::String(mpb->name.c_str());
		tmp += "を関連付けしました";
#ifdef CREATE_EXE_EXCELRTC
		System::Windows::Forms::MessageBox::Show(tmp);
#endif
	}
	else if(GetInPort(mpb->name) != NULL)
	{
#ifdef CREATE_EXE_EXCELRTC
		System::Windows::Forms::MessageBox::Show("インポートの名前が正しくありません");
#endif
	}
	else
	{
#ifdef CREATE_EXE_EXCELRTC
		System::Windows::Forms::MessageBox::Show("アウトポートを選択してください");
#endif
	}

	Save();
}
void ExcelControl::DetachPort(MyPortBase *mpb, std::string n)
{
	MyPortBase *ip = GetInPort(n);

	if(ip != NULL && GetInPort(mpb->name) == NULL)
	{
		System::String^ tmp = gcnew System::String(n.c_str());
		tmp += "と";
		tmp += gcnew System::String(mpb->name.c_str());
		tmp += "の関連付けを解除しました";
#ifdef CREATE_EXE_EXCELRTC
		System::Windows::Forms::MessageBox::Show(tmp);
#endif

	
		std::vector<std::string>::iterator end_it = remove( ip->attachPort.begin(), ip->attachPort.end(), mpb->name );
		ip->attachPort.erase( end_it, ip->attachPort.end() );


		std::vector<std::string>::iterator end_it2 = remove( mpb->attachPort.begin(), mpb->attachPort.end(), n );
		mpb->attachPort.erase( end_it2, mpb->attachPort.end() );
		

	}
	else if(GetInPort(mpb->name) != NULL)
	{
#ifdef CREATE_EXE_EXCELRTC
		System::Windows::Forms::MessageBox::Show("アウトポートを選択してください");
#endif
	}
	else
	{
#ifdef CREATE_EXE_EXCELRTC
		System::Windows::Forms::MessageBox::Show("インポートの名前が正しくありません");
#endif
	}

	Save();
}

MyPortBase *ExcelControl::GetInPort(std::string n)
{
	for(int i=0;i < InPorts.size();i++)
	{
		if(InPorts[i]->name == n)
		{
			return InPorts[i];
		}
	}
	return NULL;
}

MyPortBase *ExcelControl::GetOutPort(std::string n)
{
	for(int i=0;i < OutPorts.size();i++)
	{
		if(OutPorts[i]->name == n)
		{
			return OutPorts[i];
		}
	}
	return NULL;
}
MyPortBase *ExcelControl::GetConfInPort(std::string n)
{
	for(int i=0;i < ConfInPorts.size();i++)
	{
		if(ConfInPorts[i]->name == n)
		{
			return ConfInPorts[i];
		}
	}
	
	return NULL;
}

MyPortBase *ExcelControl::GetConfOutPort(std::string n)
{
	
	for(int i=0;i < ConfOutPorts.size();i++)
	{
		if(ConfOutPorts[i]->name == n)
		{
			return ConfOutPorts[i];
		}
	}
	return NULL;
}



void ExcelControl::DeleteAllPort()
{
	for(int i=0;i < InPorts.size();i++)
	{
		InPorts[i]->pb->disconnect_all();
		removePort(*InPorts[i]->pb);

		InPorts[i]->mop->mpb = NULL;
		delete InPorts[i];
		

	}
	for(int i=0;i < OutPorts.size();i++)
	{
		OutPorts[i]->pb->disconnect_all();
		removePort(*OutPorts[i]->pb);

		OutPorts[i]->mop->mpb = NULL;
		delete OutPorts[i];
	}
	InPorts.clear();
	OutPorts.clear();
}
void ExcelControl::Save()
{
	//update_cellName();
	std::vector<MyPortBase*>tf;
	for(int i=0;i < InPorts.size();i++)
	{
		tf.push_back(InPorts[i]);
	}
	for(int i=0;i < OutPorts.size();i++)
	{
		tf.push_back(OutPorts[i]);
	}

	std::vector<std::string> v;
	for(int i=0;i < tf.size();i++)
	{
		std::string sw = "";
		for(int j=0;j < tf[i]->mop->buff.size();j++)
		{
			sw += tf[i]->mop->buff[j];
			if(j != tf[i]->mop->buff.size()-1)
			{
				sw += "/";
			}
		}
		sw += "#";
		stringstream ss;
		ss << tf[i]->col << "#" << tf[i]->low << "#" << tf[i]->sheetName<< "#" << tf[i]->length << "#" << tf[i]->state << "#";
		for(int j=0;j < tf[i]->attachPort.size();j++)
		{
			if(j != 0)
				ss << "/";
			ss << tf[i]->attachPort[j];
			
		}
		sw += ss.str();
		v.push_back(sw);
	}

	myExcel::Obj->SaveRTC(v);
}
void ExcelControl::Load()
{
	if(myExcel::Obj)
	{
		std::vector<std::string> lw = myExcel::Obj->LoadRTC();
	

	

		if(lw.size() > 0)
		{
		
			for(int i=0;i < lw.size();i++){
			
				if(lw[i] != ""){
				
				
					vector<string> confs = split(lw[i], "#");
				
					if(confs.size() > 4)
					{
					
						vector<string> mpath = split(confs[0], "/");
						int col = string2binary<int>(confs[1],10);
						std::string low = confs[2];
						string sheetname = confs[3];
						std::string len = confs[4];
						bool mstate = true;
						if(confs[5] == "0")
							mstate = false;

						if(mpath.size() > 1 && i==0)
						{
							GetRTCTree(mpath[0]);
						}
						MyPortBase *mpb = SetDPort(mpath, col, low, sheetname, len, mstate, false);

						vector<string> atrtc = split(confs[6], "/");
					
						for(int j=0;j < atrtc.size();j++)
						{
							if(atrtc[j] != "")
							{
								mpb->attachPort.push_back(atrtc[j]);
							}
						}
					}
				
				}
			}
		}
	}
	
}

void ExcelControl::ResetPort(MyPortBase* mpb)
{
	mpb->num = 0;
}
void ExcelControl::ResetAllPort()
{
	for(int i=0;i < InPorts.size();i++)
	{
				
		InPorts[i]->num = 0;
		
		

	}
	for(int i=0;i < OutPorts.size();i++)
	{
		if(OutPorts[i]->num > 0)
			myExcel::Obj->ResetCellColor(OutPorts[i]->col+OutPorts[i]->num-1, OutPorts[i]->low, OutPorts[i]->sheetName, OutPorts[i]->length);

		OutPorts[i]->num = 0;
		
	}


	for(int i=0;i < ConfInPorts.size();i++)
	{
				
		ConfInPorts[i]->num = 0;
		
		

	}
	for(int i=0;i < ConfOutPorts.size();i++)
	{
		if(ConfOutPorts[i]->num > 0)
			myExcel::Obj->ResetCellColor(ConfOutPorts[i]->col+ConfOutPorts[i]->num-1, ConfOutPorts[i]->low, ConfOutPorts[i]->sheetName, ConfOutPorts[i]->length);

		ConfOutPorts[i]->num = 0;
		
	}
}

MyPortBase* ExcelControl::CreatePort(OtherPort &op, int c, std::string l, std::string sn, std::string leng, bool mstate)
{
	
	
	
	string DataType = NVUtil::toString(op.pb->get_port_profile()->properties,"dataport.data_type");

	
	
	
	
	string tdt = split(DataType, ":")[1];

	tdt = Replace(tdt, "RTC/", "");

	MyPortBase* mpb = NULL;
	

	if(tdt == "TimedDouble")
	{
		
		mpb = crPort<RTC::TimedDouble, double>(op, tdt, c, l, sn, leng, mstate);
	}
	else if(tdt == "TimedLong")
	{
		mpb = crPort<RTC::TimedLong, long>(op, tdt, c, l, sn, leng, mstate);
	}
	

	else if(tdt == "TimedFloat")
	{
		mpb = crPort<RTC::TimedFloat, float>(op, tdt, c, l, sn, leng, mstate);
	}
	else if(tdt == "TimedShort")
	{
		
		mpb = crPort<RTC::TimedShort, short>(op, tdt, c, l, sn, leng, mstate);
	}
	else if(tdt == "TimedULong")
	{
		mpb = crPort<RTC::TimedULong, long>(op, tdt, c, l, sn, leng, mstate);
	}
	
	else if(tdt == "TimedUShort")
	{
		mpb = crPort<RTC::TimedUShort, short>(op, tdt, c, l, sn, leng, mstate);
	}
	else if(tdt == "TimedChar")
	{
		mpb = crPort<RTC::TimedChar, char>(op, tdt, c, l, sn, leng, mstate);
	}
	else if(tdt == "TimedWChar")
	{
		//mpb = crPort<RTC::TimedWChar, char>(op, tdt, c, l, sn, leng, mstate);
	}
	else if(tdt == "TimedBoolean")
	{
		mpb = crPort<RTC::TimedBoolean, bool>(op, tdt, c, l, sn, leng, mstate);
	}
	else if(tdt == "TimedOctet")
	{
		
		mpb = crPort<RTC::TimedOctet, char>(op, tdt, c, l, sn, leng, mstate);
	}
	else if(tdt == "TimedString")
	{
		
		//mpb = crPort<RTC::TimedString, char*>(op, tdt, c, l, sn, leng, mstate);
	}
	else if(tdt == "TimedWString")
	{
		
		//mpb = crPort<RTC::TimedWString, char*>(op, tdt, c, l, sn, leng, mstate);
	}
	

	else if(tdt == "TimedDoubleSeq")
	{
		
		mpb = crPortSeq<RTC::TimedDoubleSeq, double>(op, tdt, c, l, sn, leng, mstate);
	}
	else if(tdt == "TimedLongSeq")
	{
		mpb = crPortSeq<RTC::TimedLongSeq, long>(op, tdt, c, l, sn, leng, mstate);
	}

	else if(tdt == "TimedFloatSeq")
	{
		mpb = crPortSeq<RTC::TimedFloatSeq, float>(op, tdt, c, l, sn, leng, mstate);
	}
	
	else if(tdt == "TimedShortSeq")
	{
		mpb = crPortSeq<RTC::TimedShortSeq, short>(op, tdt, c, l, sn, leng, mstate);
	}
	else if(tdt == "TimedULongSeq")
	{
		
		mpb = crPortSeq<RTC::TimedULongSeq, long>(op, tdt, c, l, sn, leng, mstate);
	}
	else if(tdt == "TimedUShortSeq")
	{
		mpb = crPortSeq<RTC::TimedUShortSeq, short>(op, tdt, c, l, sn, leng, mstate);
	}
	else if(tdt == "TimedCharSeq")
	{
		mpb = crPortSeq<RTC::TimedCharSeq, char>(op, tdt, c, l, sn, leng, mstate);
	}
	else if(tdt == "TimedWCharSeq")
	{
		//mpb = crPortSeq<RTC::TimedWCharSeq, char>(op, tdt, c, l, sn, leng, mstate);
	}
	
	else if(tdt == "TimedOctetSeq")
	{
		mpb = crPortSeq<RTC::TimedOctetSeq, char>(op, tdt, c, l, sn, leng, mstate);
	}

	else if(tdt == "TimedStringSeq")
	{
		//mpb = crPortSeq<RTC::TimedStringSeq, char*>(op, tdt, c, l, sn, leng, mstate);
	}

	else if(tdt == "TimedWStringSeq")
	{
		//mpb = crPortSeq<RTC::TimedWStringSeq, char*>(op, tdt, c, l, sn, leng, mstate);
	}
	
	
	
	

	return mpb;
}


extern "C"
{
 
  void ExcelRTCInit(RTC::Manager* manager)
  {
    coil::Properties profile(excelrtc_spec);
    manager->registerFactory(profile,
                             RTC::Create<ExcelControl>,
                             RTC::Delete<ExcelControl>);
  }
  
};