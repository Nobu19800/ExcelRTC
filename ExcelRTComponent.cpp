#include "stdafx.h"
#include "ExcelRTComponent.h"

#include <iostream>
#include <fstream>
#include "MyObject.h"



 
using namespace std;


ExcelRTComponent *tertc;

static const char* excelrtc_spec[] =
  {
    "implementation_id", "ExcelRTC",
    "type_name",         "Excel Component",
    "description",       "ModuleDescription",
    "version",           "0.0.2",
    "vendor",            "Miyamoto Nobuhiko",
    "category",          "Category",
    "activity_type",     "PERIODIC",
    "kind",              "DataFlowComponent",
    "max_instance",      "1",
    "language",          "C++",
    "lang_type",         "compile",
	"conf.default.file_path", "NewFile",
	"conf.__widget__.file_path", "text",
	""
  };


//コンフィギュレーションパラメータが更新されたときのコールバック
class MyConfigUpdateParam
    : public RTC::ConfigurationSetListener
{
public:
    MyConfigUpdateParam(ExcelRTComponent *e_rtc)
    {
		m_rtc = e_rtc;
    }
    void operator()(const coil::Properties& config_set)
	{
		
		m_rtc->ConfigUpdate();
		
    }
	ExcelRTComponent *m_rtc;

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





ExcelRTComponent::ExcelRTComponent(RTC::Manager* manager)
  : RTC::DataFlowComponentBase(manager),
  m_DataBasePort("DataBase")

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

	

	
}

ExcelRTComponent::~ExcelRTComponent()
{
}

TreeObject* ExcelRTComponent::GetRTCTree(string IP_adress)
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
		System::Windows::Forms::MessageBox::Show("ネーミングサービスへの接続に失敗しました");
		return NULL;
	}
	
	//SetTree(to);
	
	
	
	
	
}

RTC::ReturnCode_t ExcelRTComponent::onInitialize()
{
  
  this->addConfigurationSetListener(ON_SET_CONFIG_SET, new MyConfigUpdateParam(this));

    
  bindParameter("file_path", file_path, "NewFile");

  
  m_DataBasePort.registerProvider("database", "DataBase::mDataBase", m_database);

  addPort(m_DataBasePort);

 
  
  return RTC::RTC_OK;
}


RTC::ReturnCode_t ExcelRTComponent::onDeactivated(RTC::UniqueId ec_id)
{
	ResetAllPort();

	return RTC::RTC_OK;
}







RTC::ReturnCode_t ExcelRTComponent::onExecute(RTC::UniqueId ec_id)
{
	//ExcelRTC::Form1::m_form->m_excel->xlApplication->ScreenUpdating = false;

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
	//ExcelRTC::Form1::m_form->m_excel->xlApplication->ScreenUpdating = true;
	

	for(int i=0;i < OutPorts.size();i++)
	{
		if(OutPorts[i]->attachPort.size() > 0)
		{
			bool in = true;
			for(int j=0;j < OutPorts[i]->attachPort.size();j++)
			{
				MyPortBase *mpb = GetInPort(OutPorts[i]->attachPort[j]);
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
				for(int j=0;j < OutPorts[i]->attachPort.size();j++)
				{
					MyPortBase *mpb = GetInPort(OutPorts[i]->attachPort[j]);
					if(mpb != NULL)
					{
						mpb->PutData(false);
					}
				}
				OutPorts[i]->PutData(false);
			}
			
			

		}
	}
	
	

  return RTC::RTC_OK;
}


void ExcelRTComponent::DelDPort(std::vector<std::string> pt)
{
	for(int i=0;i < rtclist.size();i++)
	{
		if(rtclist[i].buff == pt && rtclist[i].mpb != NULL)
		{
			DeleteOtherPort(rtclist[i]);
		}

	}
}

void ExcelRTComponent::ConfigUpdate()
{
	this->m_configsets.update("default","file_path");
	std::string sfn = Replace(file_path, "/", "\\");
	System::String ^tfn = gcnew System::String(sfn.c_str());
	//System::Console::WriteLine(tfn);
	
	if(sfn ==  "NewFile")
	{
		ExcelRTC::Form1::m_form->m_excel->Open("");
	}
	else if(ExcelRTC::Form1::m_form->m_excel->filename != tfn)
	{
		
		ExcelRTC::Form1::m_form->m_excel->Open(tfn);
		 
		
	}
	Load();
}

MyPortBase* ExcelRTComponent::GetDPort(std::vector<std::string> pt)
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

MyPortBase* ExcelRTComponent::SetDPort(std::vector<std::string> pt, int c, std::string l, std::string sn, std::string leng, bool mstate)
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

			System::Windows::Forms::MessageBox::Show(tmp);

			return CreatePort(rtclist[i], c, l, sn, leng, mstate);
		}
		else if(rtclist[i].buff == pt && rtclist[i].mpb != NULL)
		{
			rtclist[i].mpb->SetExcelParam(c, l, sn, leng, mstate);
			return rtclist[i].mpb;
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

void ExcelRTComponent::DeleteOtherPort(OtherPort &op)
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

	System::Windows::Forms::MessageBox::Show("削除しました");

	
}

void ExcelRTComponent::AttachPort(MyPortBase *mpb, std::string n)
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
		System::Windows::Forms::MessageBox::Show(tmp);
	}
	else if(GetInPort(mpb->name) != NULL)
	{
		System::Windows::Forms::MessageBox::Show("インポートの名前が正しくありません");
	}
	else
	{
		System::Windows::Forms::MessageBox::Show("アウトポートを選択してください");
	}

	Save();
}
void ExcelRTComponent::DetachPort(MyPortBase *mpb, std::string n)
{
	MyPortBase *ip = GetInPort(n);

	if(ip != NULL && GetInPort(mpb->name) == NULL)
	{
		System::String^ tmp = gcnew System::String(n.c_str());
		tmp += "と";
		tmp += gcnew System::String(mpb->name.c_str());
		tmp += "の関連付けを解除しました";
		System::Windows::Forms::MessageBox::Show(tmp);

	
		std::vector<std::string>::iterator end_it = remove( ip->attachPort.begin(), ip->attachPort.end(), mpb->name );
		ip->attachPort.erase( end_it, ip->attachPort.end() );


		std::vector<std::string>::iterator end_it2 = remove( mpb->attachPort.begin(), mpb->attachPort.end(), n );
		mpb->attachPort.erase( end_it2, mpb->attachPort.end() );
		

	}
	else if(GetInPort(mpb->name) != NULL)
	{
		System::Windows::Forms::MessageBox::Show("アウトポートを選択してください");
	}
	else
	{
		System::Windows::Forms::MessageBox::Show("インポートの名前が正しくありません");
	}

	Save();
}

MyPortBase *ExcelRTComponent::GetInPort(std::string n)
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

MyPortBase *ExcelRTComponent::GetOutPort(std::string n)
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

void ExcelRTComponent::DeleteAllPort()
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
void ExcelRTComponent::Save()
{
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

	ExcelRTC::Form1::m_form->m_excel->SaveRTC(v);
}
void ExcelRTComponent::Load()
{
	std::vector<std::string> lw = ExcelRTC::Form1::m_form->m_excel->LoadRTC();
	

	

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
					MyPortBase *mpb = SetDPort(mpath, col, low, sheetname, len, mstate);

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

void ExcelRTComponent::ResetPort(MyPortBase* mpb)
{
	mpb->num = 0;
}
void ExcelRTComponent::ResetAllPort()
{
	for(int i=0;i < InPorts.size();i++)
	{
				
		InPorts[i]->num = 0;
		
		

	}
	for(int i=0;i < OutPorts.size();i++)
	{
		if(OutPorts[i]->num > 0)
			ExcelRTC::Form1::m_form->m_excel->ResetCellColor(OutPorts[i]->col+OutPorts[i]->num-1, OutPorts[i]->low, OutPorts[i]->sheetName, OutPorts[i]->length);

		OutPorts[i]->num = 0;
		
	}
}

MyPortBase* ExcelRTComponent::CreatePort(OtherPort &op, int c, std::string l, std::string sn, std::string leng, bool mstate)
{
	
	
	
	string DataType = NVUtil::toString(op.pb->get_port_profile()->properties,"dataport.data_type");

	
	
	
	
	string tdt = split(DataType, ":")[1];

	MyPortBase* mpb = NULL;
	

	if(tdt == "RTC/TimedDouble")
	{
		
		mpb = crPort<RTC::TimedDouble, double>(op, tdt, op, c, l, sn, leng, mstate);
	}
	else if(tdt == "RTC/TimedLong")
	{
		mpb = crPort<RTC::TimedLong, long>(op, tdt, op, c, l, sn, leng, mstate);
	}
	

	else if(tdt == "RTC/TimedFloat")
	{
		mpb = crPort<RTC::TimedFloat, float>(op, tdt, op, c, l, sn, leng, mstate);
	}
	else if(tdt == "RTC/TimedShort")
	{
		
		mpb = crPort<RTC::TimedShort, short>(op, tdt, op, c, l, sn, leng, mstate);
	}
	else if(tdt == "RTC/TimedULong")
	{
		mpb = crPort<RTC::TimedULong, long>(op, tdt, op, c, l, sn, leng, mstate);
	}
	
	else if(tdt == "RTC/TimedUShort")
	{
		mpb = crPort<RTC::TimedUShort, short>(op, tdt, op, c, l, sn, leng, mstate);
	}
	else if(tdt == "RTC/TimedChar")
	{
		mpb = crPort<RTC::TimedChar, char>(op, tdt, op, c, l, sn, leng, mstate);
	}
	else if(tdt == "RTC/TimedWChar")
	{
		//mpb = crPort<RTC::TimedWChar, char>(op, tdt, op, c, l, sn, leng, mstate);
	}
	else if(tdt == "RTC/TimedBoolean")
	{
		mpb = crPort<RTC::TimedBoolean, bool>(op, tdt, op, c, l, sn, leng, mstate);
	}
	else if(tdt == "RTC/TimedOctet")
	{
		
		mpb = crPort<RTC::TimedOctet, char>(op, tdt, op, c, l, sn, leng, mstate);
	}
	else if(tdt == "RTC/TimedString")
	{
		
		//mpb = crPort<RTC::TimedString, char*>(op, tdt, op, c, l, sn, leng, mstate);
	}
	else if(tdt == "RTC/TimedWString")
	{
		
		//mpb = crPort<RTC::TimedWString, char*>(op, tdt, op, c, l, sn, leng, mstate);
	}
	

	else if(tdt == "RTC/TimedDoubleSeq")
	{
		
		mpb = crPortSeq<RTC::TimedDoubleSeq, double>(op, tdt, op, c, l, sn, leng, mstate);
	}
	else if(tdt == "RTC/TimedLongSeq")
	{
		mpb = crPortSeq<RTC::TimedLongSeq, long>(op, tdt, op, c, l, sn, leng, mstate);
	}

	else if(tdt == "RTC/TimedFloatSeq")
	{
		mpb = crPortSeq<RTC::TimedFloatSeq, float>(op, tdt, op, c, l, sn, leng, mstate);
	}
	
	else if(tdt == "RTC/TimedShortSeq")
	{
		mpb = crPortSeq<RTC::TimedShortSeq, short>(op, tdt, op, c, l, sn, leng, mstate);
	}
	else if(tdt == "RTC/TimedULongSeq")
	{
		
		mpb = crPortSeq<RTC::TimedULongSeq, long>(op, tdt, op, c, l, sn, leng, mstate);
	}
	else if(tdt == "RTC/TimedUShortSeq")
	{
		mpb = crPortSeq<RTC::TimedUShortSeq, short>(op, tdt, op, c, l, sn, leng, mstate);
	}
	else if(tdt == "RTC/TimedCharSeq")
	{
		mpb = crPortSeq<RTC::TimedCharSeq, char>(op, tdt, op, c, l, sn, leng, mstate);
	}
	else if(tdt == "RTC/TimedWCharSeq")
	{
		//mpb = crPortSeq<RTC::TimedWCharSeq, char>(op, tdt, op, c, l, sn, leng, mstate);
	}
	
	else if(tdt == "RTC/TimedOctetSeq")
	{
		mpb = crPortSeq<RTC::TimedOctetSeq, char>(op, tdt, op, c, l, sn, leng, mstate);
	}

	else if(tdt == "RTC/TimedStringSeq")
	{
		//mpb = crPortSeq<RTC::TimedStringSeq, char*>(op, tdt, op, c, l, sn, leng, mstate);
	}

	else if(tdt == "RTC/TimedWStringSeq")
	{
		//mpb = crPortSeq<RTC::TimedWStringSeq, char*>(op, tdt, op, c, l, sn, leng, mstate);
	}
	
	
	
	

	return mpb;
}


extern "C"
{
 
  void ExcelRTCInit(RTC::Manager* manager)
  {
    coil::Properties profile(excelrtc_spec);
    manager->registerFactory(profile,
                             RTC::Create<ExcelRTComponent>,
                             RTC::Delete<ExcelRTComponent>);
  }
  
};