// -*-C++-*-
/*!
 * @file  ExcelControl.cpp
 * @brief ExcelControlComponent
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
#include "SubFunction.h"



 
using namespace std;


ExcelControl *tertc;

static const char* excelrtc_spec[] =
  {
    "implementation_id", "ExcelControl",
    "type_name",         "ExcelControl",
    "description",       "Excel Control Component",
    "version",           "0.0.2",
    "vendor",            "Miyamoto Nobuhiko",
    "category",          "Office",
    "activity_type",     "PERIODIC",
    "kind",              "DataFlowComponent",
    "max_instance",      "1",
    "language",          "C++",
    "lang_type",         "compile",
	"conf.default.file_path", "NewFile",
	"conf.default.actionLock", "0",
    "conf.default.red", "255",
    "conf.default.green", "255",
    "conf.default.blue", "0",
	"conf.dataport0.port_type", "DataInPort",
	"conf.dataport0.data_type", "TimedFloat",
	"conf.dataport0.column", "1",
	"conf.dataport0.start_row", "A",
	"conf.dataport0.end_row", "A",
	"conf.dataport0.sheetname", "Sheet1",
	"conf.dataport0.c_move", "1",
    "conf.dataport0.attach_Port", "None",
	"conf.__widget__.actionLock", "radio",
    "conf.__widget__.red", "spin",
    "conf.__widget__.green", "spin",
    "conf.__widget__.blue", "spin",
	"conf.__widget__.file_path", "text",
	"conf.__widget__.port_type", "radio",
	"conf.__widget__.column", "spin",
	"conf.__widget__.start_row", "text",
	"conf.__widget__.end_row", "text",
	"conf.__widget__.sheetname", "text",
	"conf.__widget__.data_type", "radio",
	"conf.__widget__.c_move", "radio",
    "conf.__widget__.attach_Port", "text",
	"conf.__constraints__.actionLock", "(0,1)",
    "conf.__constraints__.red", "0<=x<=255",
    "conf.__constraints__.green", "0<=x<=255",
    "conf.__constraints__.blue", "0<=x<=255",
    "conf.__constraints__.column", "1<=x<=1000",
	"conf.__constraints__.column", "1<=x<=1000",
	"conf.__constraints__.port_type", "(DataInPort,DataOutPort)",
	"conf.__constraints__.data_type", "(TimedDouble,TimedLong,TimedFloat,TimedShort,TimedULong,TimedUShort,TimedChar,TimedWChar,TimedBoolean,TimedOctet,TimedString,TimedWString,TimedDoubleSeq,TimedLongSeq,TimedFloatSeq,TimedShortSeq,TimedULongSeq,TimedUShortSeq,TimedCharSeq,TimedWCharSeq,TimedOctetSeq,TimedStringSeq,TimedWStringSeq)",
	"conf.__constraints__.c_move", "(0,1)",
	""
  };







/*void SetTree(TreeObject *to)
{
	
	
	//System::String^ a = gcnew System::String(to->name.c_str());
	//System::Windows::Forms::MessageBox::Show(a);

	

	for(int i=0;i < to->to.size();i++)
	{
		SetTree(to->to[i]);
	}
}*/




/**
*@brief Excelを操作するRTコンポーネントのコンストラクタ
* @param manager マネージャオブジェクト
*/
ExcelControl::ExcelControl(RTC::Manager* manager)
  : RTC::DataFlowComponentBase(manager),
  m_SpreadSheetPort("SpreadSheet"),
  m_spreadsheet()

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
		createPort(rtclist[i]);
	}*/
	

	
	

	
}

/**
*@brief Excelを操作するRTコンポーネントのデストラクタ
*/
ExcelControl::~ExcelControl()
{
}

/**
*@brief ファイルを取得する関数
* @return ファイル名
*/
std::string ExcelControl::getFileName()
{

	return file_path;
}

/**
*@brief RTCのデータポートのツリーを取得する関数
* @param IP_adress ネームサーバの名前
* @return ツリーオブジェクト
*/
TreeObject* ExcelControl::getRTCTree(string IP_adress)
{
	try
	{
		RTC::CorbaNaming namingserver(m_manager->getORB(), IP_adress.c_str());
		rtclist.clear();
		deleteAllPort();
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

/**
*@brief 初期化処理用コールバック関数
* @return RTC::ReturnCode_t
*/
RTC::ReturnCode_t ExcelControl::onInitialize()
{
	//ExcelObject::Obj = gcnew ExcelObject();
  
  


  bindParameter("file_path", file_path, "NewFile");
  bindParameter("data_type", conf_data_type, "TimedFloat");
  bindParameter("port_type", conf_port_type, "DataInPort");
  bindParameter("column", conf_column, "1");
  bindParameter("start_row", conf_start_row, "A");
  bindParameter("end_row", conf_end_row, "A");
  bindParameter("sheetname", conf_sheetname, "Sheet1");
  bindParameter("actionLock", actionLock, "0");
  bindParameter("red", red, "255");
  bindParameter("green", green, "255");
  bindParameter("blue", blue, "0");
  bindParameter("c_move", c_move, "1");
  bindParameter("attach_Port", attach_Port, "None");

  this->m_configsets.update("default", "file_path");
  et = new ExcelTask(this);
  et->activate();

  m_SpreadSheetPort.registerProvider("spreadsheet", "SpreadSheet::mSpreadSheet", m_spreadsheet);

  addPort(m_SpreadSheetPort);
  
  

  /*std::string filePath = "";
  coil::Properties& prop(::RTC::Manager::instance().getConfig());
  getProperty(prop, "excel.filename", filePath);
  setFilePath(filePath);*/


  this->addConfigurationSetListener(ON_SET_CONFIG_SET, new ExcelConfigUpdateParam(this));
  

  
  

  
  
 
  

  return RTC::RTC_OK;
}


/**
*@brief ファイル名のコンフィギュレーションパラメータ変更の関数
* @param FP ファイル名
*/
void ExcelControl::setFilePath(std::string FP)
{
	
	

	//std::string file_key = "file_path";

	//file_path = FP;
	
	

	coil::Properties file_confSet("default");// = this->m_configsets.getConfigurationSet("default");
	file_confSet.setProperty("file_path", FP.c_str());
	this->m_configsets.setConfigurationSetValues(file_confSet);
	this->m_configsets.activateConfigurationSet("default");
	//bindParameter("file_path", file_path, "NewFile");


	this->m_configsets.update("default", "file_path");
	
}

/**
*@brief 不活性化時のコールバック関数
* @param ec_id target ExecutionContext Id
* @return RTC::ReturnCode_t
*/
RTC::ReturnCode_t ExcelControl::onDeactivated(RTC::UniqueId ec_id)
{
	resetAllPort();

	return RTC::RTC_OK;
}

/**
*@brief 終了処理のコールバック関数
* @return RTC::ReturnCode_t
*/
RTC::ReturnCode_t ExcelControl::onFinalize()
{
  ExcelObject::Obj->Close();
  return RTC::RTC_OK;
}

/**
*@brief セルの名前を記入
*/
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



/**
*@brief 周期処理用コールバック関数
* @param ec_id target ExecutionContext Id
* @return RTC::ReturnCode_t
*/
RTC::ReturnCode_t ExcelControl::onExecute(RTC::UniqueId ec_id)
{
	if(actionLock == 1)
	{
		ExcelObject::Obj->xlApplication->ScreenUpdating = false;
		_mutex.lock();
	}

	ExcelObject::Obj->setColor(red, green, blue);

	for(int i=0;i < ConfInPorts.size();i++)
	{
		if(ConfInPorts[i]->attachPort.size() == 0)
		{
			ConfInPorts[i]->putData(true);
		}
		
	}

	for(int i=0;i < ConfOutPorts.size();i++)
	{
		if(ConfOutPorts[i]->attachPort.size() == 0)
		{
			ConfOutPorts[i]->putData(true);
		}

	}


	for(int i=0;i < InPorts.size();i++)
	{
		if(InPorts[i]->attachPort.size() == 0)
		{
			InPorts[i]->putData(true);
			
		}
		
		
	}

	for(int i=0;i < OutPorts.size();i++)
	{
		if(OutPorts[i]->attachPort.size() == 0)
		{
			OutPorts[i]->putData(true);
			
		}
	}

	if(actionLock == 1)
	{
		ExcelObject::Obj->xlApplication->ScreenUpdating = true;
		_mutex.unlock();
	}
	

	
	
	

  return RTC::RTC_OK;
}

/**
*@brief 関連付けしたデータポートの処理
* @param ip インポート
*/
void ExcelControl::updateAPort(ExcelPortBase* ip)
{
	for(int i=0;i < ip->attachPort.size();i++)
	{
		ExcelPortBase *op = getOutPort(ip->attachPort[i]);
		if(op == NULL)
		{
			op = getConfOutPort(ip->attachPort[i]);
		}


		if(op)
		{
			if(op->attachPort.size() > 0)
			{
				bool in = true;
				for(int j=0;j < op->attachPort.size();j++)
				{
					ExcelPortBase *mpb = getInPort(op->attachPort[j]);
					if(mpb == NULL)
					{
						mpb = getConfInPort(op->attachPort[j]);
					}

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
						ExcelPortBase *mpb = getInPort(op->attachPort[j]);
						if(mpb == NULL)
						{
							mpb = getConfInPort(op->attachPort[j]);
						}

						if(mpb != NULL)
						{
							mpb->putData(false);
						}
					}
					op->putData(false);

					_mutex.unlock();
				}
			
			

			}
		}
	}
}

/**
*@brief データポートを削除する関数
* @param pt 削除するデータポートのパス
*/
void ExcelControl::delDPort(std::vector<std::string> pt)
{
	for(int i=0;i < rtclist.size();i++)
	{
		if(rtclist[i].buff == pt && rtclist[i].mpb != NULL)
		{
			deleteOtherPort(rtclist[i]);
		}

	}
}

/**
*@brief コンフィギュレーションパラメータが変更されたときに呼び出される関数
*/
void ExcelControl::configUpdate()
{

	this->m_configsets.update("default", "file_path");
	std::string sfn = file_path;
	coil::replaceString(sfn, "/", "\\");
	
	System::String ^tfn = gcnew System::String(sfn.c_str());
	//System::Console::WriteLine(tfn);
	
	if(sfn ==  "NewFile")
	{
		ExcelObject::Obj->Open("");
		load();
	}
	else if(ExcelObject::Obj->filename != tfn)
	{
		
		ExcelObject::Obj->Open(tfn);
		load();
		 
		
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
			this->m_configsets.update(dn.c_str());

			std::string tdt = "DataInPort";

			if(conf_column > 0 && conf_start_row.size() > 0)
			{

				ExcelPortBase* tmp = getConfInPort(dn);
				if(tmp == NULL)
				{
					tmp = getConfOutPort(dn);
					tdt = "DataOutPort";
				}
				bool t_c_move = true;
				if(c_move == 0)
				{
					t_c_move = false;
				}
				std::vector<std::string> t_attach_Port;
				std::istringstream stream( attach_Port );
				std::string token;
				while( std::getline( stream, token, ',' ) )
				{
					if(token != "None" && token != "")
					{
						t_attach_Port.push_back(token);
					}
				}
				if(tmp && tmp->data_type == conf_data_type && conf_port_type == tdt)
				{
					tmp->setParam(conf_column, conf_start_row, conf_sheetname, conf_end_row, t_c_move);
					tmp->attachPort = t_attach_Port;
				}
				else
				{
					if(tmp)
					{
						tmp->pb->disconnect_all();
						removePort(*tmp->pb);
						
						std::vector<ExcelPortBase*>::iterator end_it = remove( ConfInPorts.begin(), ConfInPorts.end(), tmp );
						ConfInPorts.erase( end_it, ConfInPorts.end() );

						std::vector<ExcelPortBase*>::iterator end_it2 = remove( ConfOutPorts.begin(), ConfOutPorts.end(), tmp );
						ConfOutPorts.erase( end_it2, ConfOutPorts.end() );
					}
					
					
					if(conf_data_type == "TimedDouble")
					{
						ExcelPortBase* t_cp = confcrPort<RTC::TimedDouble, double>(dn, conf_port_type, conf_data_type, conf_column, conf_start_row, conf_sheetname, conf_end_row, t_c_move);
						t_cp->attachPort = t_attach_Port;
					}
					else if(conf_data_type == "TimedLong")
					{
						ExcelPortBase* t_cp = confcrPort<RTC::TimedLong, long>(dn, conf_port_type, conf_data_type, conf_column, conf_start_row, conf_sheetname, conf_end_row, t_c_move);
						t_cp->attachPort = t_attach_Port;
					}
					

					else if(conf_data_type == "TimedFloat")
					{
						
						ExcelPortBase* t_cp = confcrPort<RTC::TimedFloat, float>(dn, conf_port_type, conf_data_type, conf_column, conf_start_row, conf_sheetname, conf_end_row, t_c_move);
						t_cp->attachPort = t_attach_Port;
					}
					else if(conf_data_type == "TimedShort")
					{
						
						ExcelPortBase* t_cp = confcrPort<RTC::TimedShort, short>(dn, conf_port_type, conf_data_type, conf_column, conf_start_row, conf_sheetname, conf_end_row, t_c_move);
						t_cp->attachPort = t_attach_Port;
					}
					else if(conf_data_type == "TimedULong")
					{
						ExcelPortBase* t_cp = confcrPort<RTC::TimedULong, long>(dn, conf_port_type, conf_data_type, conf_column, conf_start_row, conf_sheetname, conf_end_row, t_c_move);
						t_cp->attachPort = t_attach_Port;
					}
					
					else if(conf_data_type == "TimedUShort")
					{
						ExcelPortBase* t_cp = confcrPort<RTC::TimedUShort, short>(dn, conf_port_type, conf_data_type, conf_column, conf_start_row, conf_sheetname, conf_end_row, t_c_move);
						t_cp->attachPort = t_attach_Port;
					}
					else if(conf_data_type == "TimedChar")
					{
						ExcelPortBase* t_cp = confcrPort<RTC::TimedChar, char>(dn, conf_port_type, conf_data_type, conf_column, conf_start_row, conf_sheetname, conf_end_row, t_c_move);
						t_cp->attachPort = t_attach_Port;
					}
					else if(conf_data_type == "TimedWChar")
					{
						//ExcelPortBase* t_cp = confcrPort<RTC::TimedWChar, char>(dn, conf_port_type, conf_data_type, conf_column, conf_start_row, conf_sheetname, conf_end_row, t_c_move);
						//t_cp->attachPort = t_attach_Port;
					}
					else if(conf_data_type == "TimedBoolean")
					{
						ExcelPortBase* t_cp = confcrPort<RTC::TimedBoolean, bool>(dn, conf_port_type, conf_data_type, conf_column, conf_start_row, conf_sheetname, conf_end_row, t_c_move);
						t_cp->attachPort = t_attach_Port;
					}
					else if(conf_data_type == "TimedOctet")
					{
						
						ExcelPortBase* t_cp = confcrPort<RTC::TimedOctet, char>(dn, conf_port_type, conf_data_type, conf_column, conf_start_row, conf_sheetname, conf_end_row, t_c_move);
						t_cp->attachPort = t_attach_Port;
					}
					else if(conf_data_type == "TimedString")
					{
						
						//ExcelPortBase* t_cp = confcrPort<RTC::TimedString, char*>(dn, conf_port_type, conf_data_type, conf_column, conf_start_row, conf_sheetname, conf_end_row, t_c_move);
						//t_cp->attachPort = t_attach_Port;
					}
					else if(conf_data_type == "TimedWString")
					{
						
						//ExcelPortBase* t_cp = confcrPort<RTC::TimedWString, char*>(dn, conf_port_type, conf_data_type, conf_column, conf_start_row, conf_sheetname, conf_end_row, t_c_move);
						//t_cp->attachPort = t_attach_Port;
					}
					

					else if(conf_data_type == "TimedDoubleSeq")
					{
						
						ExcelPortBase* t_cp = confcrPortSeq<RTC::TimedDoubleSeq, double>(dn, conf_port_type, conf_data_type, conf_column, conf_start_row, conf_sheetname, conf_end_row, t_c_move);
						t_cp->attachPort = t_attach_Port;
					}
					else if(conf_data_type == "TimedLongSeq")
					{
						ExcelPortBase* t_cp = confcrPortSeq<RTC::TimedLongSeq, long>(dn, conf_port_type, conf_data_type, conf_column, conf_start_row, conf_sheetname, conf_end_row, t_c_move);
						t_cp->attachPort = t_attach_Port;
					}

					else if(conf_data_type == "TimedFloatSeq")
					{
						ExcelPortBase* t_cp = confcrPortSeq<RTC::TimedFloatSeq, float>(dn, conf_port_type, conf_data_type, conf_column, conf_start_row, conf_sheetname, conf_end_row, t_c_move);
						t_cp->attachPort = t_attach_Port;
					}
					
					else if(conf_data_type == "TimedShortSeq")
					{
						ExcelPortBase* t_cp = confcrPortSeq<RTC::TimedShortSeq, short>(dn, conf_port_type, conf_data_type, conf_column, conf_start_row, conf_sheetname, conf_end_row, t_c_move);
						t_cp->attachPort = t_attach_Port;
					}
					else if(conf_data_type == "TimedULongSeq")
					{
						
						ExcelPortBase* t_cp = confcrPortSeq<RTC::TimedULongSeq, long>(dn, conf_port_type, conf_data_type, conf_column, conf_start_row, conf_sheetname, conf_end_row, t_c_move);
						t_cp->attachPort = t_attach_Port;
					}
					else if(conf_data_type == "TimedUShortSeq")
					{
						ExcelPortBase* t_cp = confcrPortSeq<RTC::TimedUShortSeq, short>(dn, conf_port_type, conf_data_type, conf_column, conf_start_row, conf_sheetname, conf_end_row, t_c_move);
						t_cp->attachPort = t_attach_Port;
					}
					else if(conf_data_type == "TimedCharSeq")
					{
						ExcelPortBase* t_cp = confcrPortSeq<RTC::TimedCharSeq, char>(dn, conf_port_type, conf_data_type, conf_column, conf_start_row, conf_sheetname, conf_end_row, t_c_move);
						t_cp->attachPort = t_attach_Port;
					}
					else if(conf_data_type == "TimedWCharSeq")
					{
						//ExcelPortBase* t_cp = confcrPortSeq<RTC::TimedWCharSeq, char>(dn, conf_port_type, conf_data_type, conf_column, conf_start_row, conf_sheetname, conf_end_row, t_c_move);
						//t_cp->attachPort = t_attach_Port;
					}
					
					else if(conf_data_type == "TimedOctetSeq")
					{
						ExcelPortBase* t_cp = confcrPortSeq<RTC::TimedOctetSeq, char>(dn, conf_port_type, conf_data_type, conf_column, conf_start_row, conf_sheetname, conf_end_row, t_c_move);
						t_cp->attachPort = t_attach_Port;
					}

					else if(conf_data_type == "TimedStringSeq")
					{
						//ExcelPortBase* t_cp = confcrPortSeq<RTC::TimedStringSeq, char*>(dn, conf_port_type, conf_data_type, conf_column, conf_start_row, conf_sheetname, conf_end_row, t_c_move);
						//t_cp->attachPort = t_attach_Port;
					}

					else if(conf_data_type == "TimedWStringSeq")
					{
						//ExcelPortBase* t_cp = confcrPortSeq<RTC::TimedWStringSeq, char*>(dn, conf_port_type, conf_data_type, conf_column, conf_start_row, conf_sheetname, conf_end_row, t_c_move);
						//t_cp->attachPort = t_attach_Port;
					}
				}	
					
			}
		}
	}
	
}

/**
*@brief データポートを取得する関数
* @param pt 取得するデータポートのパス
* @return データポートオブジェクト
*/
ExcelPortBase* ExcelControl::getDPort(std::vector<std::string> pt)
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


/**
*@brief データポートの作成、値の設定を行う関数
* @param pt 接続するデータポートのパス
* @param c 列番号
* @param l 行番号
* @param sn シート名
* @param leng 行の範囲
* @param mstate 列を移動する場合はTrue
* @param msflag Trueの場合はメッセージボックスの表示とセルに名前の書き込みを行う
* @return データポートオブジェクト
*/
ExcelPortBase* ExcelControl::setDPort(std::vector<std::string> pt, int c, std::string l, std::string sn, std::string leng, bool mstate, bool msflag)
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

			ExcelPortBase* m_pb = createPort(rtclist[i], c, l, sn, leng, mstate);
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
			
			rtclist[i].mpb->setParam(c, l, sn, leng, mstate);

			ExcelPortBase* m_pb = rtclist[i].mpb;
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

/**
*@brief データポートを削除する関数
* @param op 削除するデータポート
*/
void ExcelControl::deleteOtherPort(OtherPort &op)
{
	

	op.mpb->pb->disconnect_all();
	removePort(*op.mpb->pb);

	
	std::vector<ExcelPortBase*>::iterator end_it = remove( InPorts.begin(), InPorts.end(), op.mpb );
	InPorts.erase( end_it, InPorts.end() );

	std::vector<ExcelPortBase*>::iterator end_it2 = remove( OutPorts.begin(), OutPorts.end(), op.mpb );
	OutPorts.erase( end_it2, OutPorts.end() );

	delete op.mpb;
	op.mpb = NULL;

	save();
#ifdef CREATE_EXE_EXCELRTC
	System::Windows::Forms::MessageBox::Show("削除しました");
#endif
	
}

/**
*@brief データポートを関連付ける関数
* @param mpb 関連付けるアウトポート
* @param n 関連付けるインポートの名前
*/
void ExcelControl::attachPort(ExcelPortBase *mpb, std::string n)
{
	
	ExcelPortBase *ip = getInPort(n);

	if(ip != NULL && getInPort(mpb->name) == NULL)
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
	else if(getInPort(mpb->name) != NULL)
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

	save();
}

/**
*@brief データポートの関連付けを解除する関数
* @param mpb 関連付けを解除するアウトポート
* @param n 関連付けを解除するインポートの名前
*/
void ExcelControl::detachPort(ExcelPortBase *mpb, std::string n)
{
	ExcelPortBase *ip = getInPort(n);

	if(ip != NULL && getInPort(mpb->name) == NULL)
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
	else if(getInPort(mpb->name) != NULL)
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

	save();
}

/**
*@brief インポートを取得する関数
* @param n インポートの名前
* @return データポートオブジェクト
*/
ExcelPortBase *ExcelControl::getInPort(std::string n)
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


/**
*@brief アウトポートを取得する関数
* @param n アウトポートの名前
* @return データポートオブジェクト
*/
ExcelPortBase *ExcelControl::getOutPort(std::string n)
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

/**
*@brief コンフィギュレーションパラメータで設定したインポートを取得する関数
* @param n インポートの名前
* @return データポートオブジェクト
*/
ExcelPortBase *ExcelControl::getConfInPort(std::string n)
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

/**
*@brief コンフィギュレーションパラメータで設定したアウトポートを取得する関数
* @param n アウトポートの名前
* @return データポートオブジェクト
*/
ExcelPortBase *ExcelControl::getConfOutPort(std::string n)
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


/**
*@brief 全てのデータポートを削除する関数
*/
void ExcelControl::deleteAllPort()
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

/**
*@brief RTCの情報を保存用シートに書き込む関数
*/
void ExcelControl::save()
{
	//update_cellName();
	std::vector<ExcelPortBase*>tf;
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

	ExcelObject::Obj->saveRTC(v);
}

/**
*@brief RTCの情報をほ場用シートより読み込む関数
*/
void ExcelControl::load()
{
	
	if(ExcelObject::Obj)
	{
		
		std::vector<std::string> lw = ExcelObject::Obj->loadRTC();
	

	

		if(lw.size() > 0)
		{
		
			for(int i=0;i < lw.size();i++){
			
				if(lw[i] != ""){
				
					
					vector<string> confs = coil::split(lw[i], "#");
				
					if(confs.size() > 4)
					{
					
						vector<string> mpath = coil::split(confs[0], "/");
						
						int col;
						coil::stringTo<int>(col,confs[1].c_str());

						std::string low = confs[2];
						string sheetname = confs[3];
						std::string len = confs[4];
						bool mstate = true;
						if(confs[5] == "0")
							mstate = false;

						if(mpath.size() > 1 && i==0)
						{
							getRTCTree(mpath[0]);
							
						}
						
						ExcelPortBase *mpb = setDPort(mpath, col, low, sheetname, len, mstate, false);
						
						if(mpb)
						{
							vector<string> atrtc = coil::split(confs[6], "/");
					
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
	
}

/**
*@brief データを書き込む列を初期化する
* @param mpb データポートオブジェクト
*/
void ExcelControl::resetPort(ExcelPortBase* mpb)
{
	mpb->num = 0;
}

/**
*@brief 全てのデータを書き込む列を初期化する
*/
void ExcelControl::resetAllPort()
{
	for(int i=0;i < InPorts.size();i++)
	{
				
		InPorts[i]->num = 0;
		
		

	}
	for(int i=0;i < OutPorts.size();i++)
	{
		if(OutPorts[i]->num > 0)
			ExcelObject::Obj->resetCellColor(OutPorts[i]->col+OutPorts[i]->num-1, OutPorts[i]->low, OutPorts[i]->sheetName, OutPorts[i]->length);

		OutPorts[i]->num = 0;
		
	}


	for(int i=0;i < ConfInPorts.size();i++)
	{
				
		ConfInPorts[i]->num = 0;
		
		

	}
	for(int i=0;i < ConfOutPorts.size();i++)
	{
		if(ConfOutPorts[i]->num > 0)
			ExcelObject::Obj->resetCellColor(ConfOutPorts[i]->col+ConfOutPorts[i]->num-1, ConfOutPorts[i]->low, ConfOutPorts[i]->sheetName, ConfOutPorts[i]->length);

		ConfOutPorts[i]->num = 0;
		
	}
}

/**
*@brief データポートを作成する関数
* @param op 接続するデータポート
* @param c 列番号
* @param l 行番号
* @param sn シート名
* @param leng 行の範囲
* @param mstate 列を移動する場合はTrue
* @return 作成したデータポート
*/
ExcelPortBase* ExcelControl::createPort(OtherPort &op, int c, std::string l, std::string sn, std::string leng, bool mstate)
{
	
	
	
	string DataType = NVUtil::toString(op.pb->get_port_profile()->properties,"dataport.data_type");

	
	
	
	
	string tdt = coil::split(DataType, ":")[1];

	
	coil::replaceString(tdt, "RTC/", "");

	ExcelPortBase* mpb = NULL;
	

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
 
  void ExcelControlInit(RTC::Manager* manager)
  {
    coil::Properties profile(excelrtc_spec);
    manager->registerFactory(profile,
                             RTC::Create<ExcelControl>,
                             RTC::Delete<ExcelControl>);
  }
  
};