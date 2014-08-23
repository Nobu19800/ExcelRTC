#ifndef EXCELRTCOMPONENT_H
#define EXCELRTCOMPONENT_H

#include <rtm/Manager.h>
#include <rtm/DataFlowComponentBase.h>
#include <rtm/CorbaPort.h>
#include <rtm/DataInPort.h>
#include <rtm/DataOutPort.h>
#include <rtm/idl/BasicDataTypeSkel.h>
#include <rtm/idl/ExtendedDataTypesSkel.h>
#include <rtm/idl/InterfaceDataTypesSkel.h>

#include <vector>
#include <string>
#include <sstream>

#include <rtm/CorbaNaming.h>
#include <rtm/RTObject.h>
#include <rtm/CorbaConsumer.h>

#include "MyObject.h"
#include "Form1.h"

using namespace RTC;

#include "DataBaseSVC_impl.h"



class ExcelRTComponent;
class MyPortBase;
class OtherPort;
class MyOnUpdate;



void ListRecursive(CosNaming::NamingContext_ptr context,std::vector<OtherPort> &rtclist,std::string &name, TreeObject *to);

int rtc_get_rtclist(RTC::CorbaNaming &naming,std::vector<OtherPort> &rtclist, TreeObject *to, std::string IP_adress);






//�C���|�[�g���f�[�^���󂯂����̃R�[���o�b�N
template <class T, class T2>
class DataListener
  : public ConnectorDataListenerT<T>
{
public:
	DataListener(MyPortBase *mp, std::vector<std::vector<T2>> &sv){
		m_port = mp;
		m_data = &sv;
	}
  virtual ~DataListener()
  {
  
  }

  virtual void operator()(const ConnectorInfo& info,
                          const T& data)
  {

	 

	  std::vector<T2> tmp;
	
	  
	  tmp.push_back(data.data);
	
	  m_port->_mutex.lock();
	  m_data->push_back(tmp);
	  m_port->_mutex.unlock();
       
  };
  MyPortBase *m_port;
  std::vector<std::vector<T2>> *m_data;
  
};


//�V�[�P���X�^�̃C���|�[�g���f�[�^���󂯎�����Ƃ��̃R�[���o�b�N
template <class T, class T2>
class SeqDataListener
  : public ConnectorDataListenerT<T>
{
public:
	SeqDataListener(MyPortBase *mp, std::vector<std::vector<T2>> &sv){
		m_port = mp;
		m_data = &sv;
	}
  virtual ~SeqDataListener()
  {
  
  }

  virtual void operator()(const ConnectorInfo& info,
                          const T& data)
  {
	std::vector<T2> tmp;
	
	for(int i=0;i < data.data.length();i++)
	{
		tmp.push_back(data.data[i]);
	}
	
	m_port->_mutex.lock();
	m_data->push_back(tmp);
	m_port->_mutex.unlock();
  };
  MyPortBase *m_port;
  std::vector<std::vector<T2>> *m_data;
  
};


//����RTC�̃f�[�^�|�[�g�̃N���X
class OtherPort
{
	public:
		OtherPort(PortService_var p, std::vector<std::string> s)
		{
			pb = p;
			buff = s;
			mpb = NULL;
		};
		PortService_var pb;
		std::vector<std::string> buff;
		MyPortBase *mpb;
};


//ExcelRTC�̃f�[�^�|�[�g�̃N���X
class MyPortBase
{
public:
	virtual ~MyPortBase()
	{
		delete pb;
	}

	//GUI�œ��͂����l��ݒ肷��֐�
	virtual void SetExcelParam(int c, std::string l, std::string sn, std::string leng, bool mstate)
	{
		col = c;
		low = l;
		sheetName = sn;
		length = leng;
		num = 0;
		state = mstate;
	}
	virtual bool isNew()
	{
		return false;
	}
	virtual void PutData(bool moption)
	{

	}

	std::string name;
	
	int col;
	std::string low;
	std::string length;
	int sheetnum;
	
	std::string sheetName;
	int num;
	OtherPort *mop;
	PortBase *pb;
	std::string data_type;
	std::vector<std::string> attachPort; 

	bool state;
	ExcelRTComponent *mexc;
	coil::Mutex _mutex;

	
	

	/*template <typename T>T GetOutPort()
	{
		return static_cast<MyOutPort<T>*>this;
	}

	template <typename T>T GetInPort()
	{
		return static_cast<MyInPort<T>*>this;
	}*/
};

//�C���|�[�g�̃N���X
template <class T, class T2>
class MyInPort : public MyPortBase
{
public:
	MyInPort(T *id, RTC::InPort<T> *ip, std::string n, std::string dt, ExcelRTComponent *m_mexc){
		In = id;
		inIn = ip;
		name = n;
		data_type = dt;
		
		col = 1;
		low = 1;
		sheetnum = 1;
		

		
		pb = ip;

		
		mexc = m_mexc;

		
		inIn->addConnectorDataListener(ON_BUFFER_WRITE, new DataListener<T,T2>(this, buff));
		
	}
	~MyInPort()
	{
		//delete inIn;
		//delete In;
	}
	//�f�[�^���擾�������𔻒肷��֐�
	virtual bool isNew()
	{
		return inIn->isNew();
	}
	//�󂯂��f�[�^���Z���ɏ������ފ֐�
	virtual void PutData(bool moption)
	{
		
		
		std::vector<std::vector<T2>> v;
		if(!moption)
		{
			
			if(inIn->isNew())
			{
				inIn->read();
				//std::cout << ip->In->data << std::endl;
				
				
				std::vector<T2> v2;
				v2.push_back(In->data);
				v.push_back(v2);

				
			}
		}
		else
		{

			_mutex.lock();
			v = buff;
			buff.clear();
			_mutex.unlock();
		}
		
		

		mexc->SetCellData<T2>(v, this);

		if(state)
			num += v.size();
		

		

		
	}

	T *In;
	RTC::InPort<T> *inIn;
	std::vector<std::vector<T2>> buff;
	
};

//�V�[�P���X�^�̃C���|�[�g�̃N���X
template <class T, class T2>
class MyInPortSeq : public MyPortBase
{
public:
	MyInPortSeq(T *id, RTC::InPort<T> *ip, std::string n, std::string dt, ExcelRTComponent *m_mexc){
		In = id;
		inIn = ip;
		name = n;
		data_type = dt;
		
		col = 1;
		low = 1;
		sheetnum = 1;
		

		
		pb = ip;

		
		mexc = m_mexc;

		
		inIn->addConnectorDataListener(ON_BUFFER_WRITE, new SeqDataListener<T,T2>(this, buff));


	}
	~MyInPortSeq()
	{
		//delete inIn;
		//delete In;
	}
	//�f�[�^���擾�������𔻒肷��֐�
	virtual bool isNew()
	{
		return inIn->isNew();
	}
	//�󂯂��f�[�^���Z���ɏ������ފ֐�
	virtual void PutData(bool moption)
	{
		std::vector<std::vector<T2>> v;
		if(!moption)
		{
			
			if(inIn->isNew())
			{
				inIn->read();
				
				std::vector<T2> v2;
				

				for(int i=0;i < In->data.length();i++)
				{
					v2.push_back(In->data[i]);
					
				}
				v.push_back(v2);

			}
		}
		else
		{

			_mutex.lock();
			v = buff;
			buff.clear();
			_mutex.unlock();
		}

		mexc->SetCellData<T2>(v, this);

		if(state)
			num += v.size();

		
				
				

			
		
	}

	T *In;
	RTC::InPort<T> *inIn;
	std::vector<std::vector<T2>> buff;
	
};


//�A�E�g�|�[�g�̃N���X
template <class T, class T2>
class MyOutPort : public MyPortBase
{
public:
	MyOutPort(T *od, RTC::OutPort<T> *op, std::string n, std::string dt, ExcelRTComponent *m_mexc){
		Out = od;
		outOut = op;
		name = n;
		data_type = dt;

		col = 1;
		low = 1;
		sheetnum = 1;
		

		
		pb = op;

		mexc = m_mexc;

		

		
	}
	~MyOutPort()
	{
		//delete outOut;
		//delete Out;
	}
	//�Z���̒l���f�[�^�|�[�g����o�͂���֐�
	virtual void PutData(bool moption)
	{
				
		std::vector<T2> v = mexc->GetCellData<T2>(this);
		
		Out->data = v[0];
		if(state)
			num += 1;
		
		
		outOut->write();
	}
	T *Out;
	RTC::OutPort<T> *outOut;
};

//�V�[�P���X�^�̃A�E�g�|�[�g�̃N���X
template <class T, class T2>
class MyOutPortSeq : public MyPortBase
{
public:
	MyOutPortSeq(T *od, RTC::OutPort<T> *op, std::string n, std::string dt, ExcelRTComponent *m_mexc){
		Out = od;
		outOut = op;
		name = n;
		data_type = dt;

		col = 1;
		low = 1;
		sheetnum = 1;
		

		
		pb = op;

		mexc = m_mexc;

		

		
	}
	~MyOutPortSeq()
	{
		//delete outOut;
		//delete Out;
	}
	//�Z���̒l���f�[�^�|�[�g����o�͂���֐�
	virtual void PutData(bool moption)
	{
		int t_l = convertStrToVal(low);
		int t_leng = convertStrToVal(length);

		int leng = t_leng-t_l+1;

		Out->data.length(leng);
		std::vector<T2> v = mexc->GetCellData<T2>(this);
		if(leng != v.size())return;
		for(int i=0;i < leng;i++)
		{
			Out->data[i] = v[i];
		}
		if(state)
			num += 1;
		
		
		outOut->write();
	}
	T *Out;
	RTC::OutPort<T> *outOut;
};

void portConnect(PortService_ptr p1, PortService_var p2);


//Excel�𑀍삷��RT�R���|�[�l���g
class ExcelRTComponent
  : public RTC::DataFlowComponentBase
{
 public:

  ExcelRTComponent(RTC::Manager* manager);

  ~ExcelRTComponent();

   //�����������p�R�[���o�b�N�֐�
   virtual RTC::ReturnCode_t onInitialize();

   //���������p�R�[���o�b�N�֐�
   virtual RTC::ReturnCode_t onExecute(RTC::UniqueId ec_id);
   
   //�s���������̃R�[���o�b�N�֐�
   virtual RTC::ReturnCode_t onDeactivated(RTC::UniqueId ec_id);

   //�f�[�^�|�[�g���쐬����֐�
   MyPortBase* CreatePort(OtherPort &op, int c, std::string l, std::string sn, std::string leng, bool mstate);
   //�f�[�^�|�[�g���폜����֐�
   void DeleteOtherPort(OtherPort &op);
   //�S�Ẵf�[�^�|�[�g���폜����֐�
   void DeleteAllPort();
   //RTC�̏���ۑ��p�V�[�g�ɏ������ފ֐�
   void Save();
   //RTC�̏����ُ�p�V�[�g���ǂݍ��ފ֐�
   void Load();
   //�f�[�^���������ޗ������������
   void ResetPort(MyPortBase* mpb);
   //�S�Ẵf�[�^���������ޗ������������
   void ResetAllPort();
   //�R���t�B�M�����[�V�����p�����[�^���ύX���ꂽ�Ƃ��ɌĂяo�����֐�
   void ConfigUpdate();

   //�Z���ɒl���������݊֐�
    template <typename T>
   void SetCellData(std::vector<std::vector<T>> dt, MyPortBase *pb)
	{
		
		ExcelRTC::Form1::m_form->m_excel->SetCellValue<T>(pb->col+pb->num, pb->low, pb->sheetName, pb->state, dt);
		
	}
   //�Z���̒l���擾����֐�
    template <typename T>
	std::vector<T> GetCellData(MyPortBase *pb)
	{
		std::vector<T> td = ExcelRTC::Form1::m_form->m_excel->GetCellValue<T>(pb->col+pb->num, pb->low, pb->sheetName, pb->length);
		
		return td;
	}
	//�V�[�P���X�^�̃f�[�^�|�[�g���쐬����֐�
	template <typename T, typename T2> MyPortBase* crPortSeq(OtherPort &op, std::string tdt, int c, std::string l, std::string sn, std::string leng, bool mstate)
   {
		string PortType = NVUtil::toString(op.pb->get_port_profile()->properties,"port.port_type");
		
	   string tname = op.buff[op.buff.size()-2];
		tname += op.buff[op.buff.size()-1];
	  
		if(PortType == "DataInPort")
		{
			T *m_out = new T();
			OutPort<T> *m_outOut = new OutPort<T>(tname.c_str(),*m_out);
			addOutPort(tname.c_str(), *m_outOut);
			
			MyPortBase *mip;
			
			mip= new MyOutPortSeq<T,T2>(m_out,m_outOut,tname,tdt,this);
			mip->mop = &op;
			
			mip->SetExcelParam(c,l,sn,leng,mstate);
			OutPorts.push_back(mip);
			portConnect(m_outOut->getPortRef(), op.pb);

			op.mpb = mip;

			return mip;
		
		}
		else if(PortType == "DataOutPort")
		{
			T *m_in = new T();
			InPort<T> *m_inIn = new InPort<T>(tname.c_str(),*m_in);

			addInPort(tname.c_str(), *m_inIn);

			MyPortBase *mip = new MyInPortSeq<T,T2>(m_in,m_inIn,tname,tdt,this);
			mip->mop = &op;
			

			mip->SetExcelParam(c,l,sn,leng,mstate);
			InPorts.push_back(mip);
			portConnect(m_inIn->getPortRef(), op.pb);

			op.mpb = mip;

			return mip;
		}
		return NULL;
	}
	//�f�[�^�|�[�g���쐬����֐�
   template <typename T, typename T2> MyPortBase* crPort(OtherPort &op, std::string tdt, int c, std::string l, std::string sn, std::string leng, bool mstate)
   {
	   string PortType = NVUtil::toString(op.pb->get_port_profile()->properties,"port.port_type");
		
	   string tname = op.buff[op.buff.size()-2];
		tname += op.buff[op.buff.size()-1];
	  
		if(PortType == "DataInPort")
		{
			T *m_out = new T();
			OutPort<T> *m_outOut = new OutPort<T>(tname.c_str(),*m_out);
			addOutPort(tname.c_str(), *m_outOut);
			
			MyPortBase *mip;
			
			mip= new MyOutPort<T,T2>(m_out,m_outOut,tname,tdt,this);
			mip->mop = &op;
			
			mip->SetExcelParam(c,l,sn,leng,mstate);
			OutPorts.push_back(mip);
			portConnect(m_outOut->getPortRef(), op.pb);

			op.mpb = mip;

			return mip;
		
		}
		else if(PortType == "DataOutPort")
		{
			T *m_in = new T();
			InPort<T> *m_inIn = new InPort<T>(tname.c_str(),*m_in);
			

			addInPort(tname.c_str(), *m_inIn);

			MyPortBase *mip = new MyInPort<T,T2>(m_in,m_inIn,tname,tdt,this);
			mip->mop = &op;
			

			mip->SetExcelParam(c,l,sn,leng,mstate);
			InPorts.push_back(mip);
			portConnect(m_inIn->getPortRef(), op.pb);

			op.mpb = mip;

			return mip;
		}
		return NULL;
   }
   //�R���t�B�M�����[�V�����p�����[�^�ɂ��V�[�P���X�^�̃f�[�^�|�[�g���쐬����֐�
   template <typename T, typename T2> MyPortBase* ConfcrPortSeq(std::string tname, std::string PortType, std::string tdt, int c, std::string l, std::string sn, std::string leng, bool mstate)
   {
		
	  
		if(PortType == "DataOutPort")
		{
			T *m_out = new T();
			OutPort<T> *m_outOut = new OutPort<T>(tname.c_str(),*m_out);
			addOutPort(tname.c_str(), *m_outOut);
			
			MyPortBase *mip;
			
			mip= new MyOutPortSeq<T,T2>(m_out,m_outOut,tname,tdt,this);
			
			mip->SetExcelParam(c,l,sn,leng,mstate);
			ConfOutPorts.push_back(mip);
			

			

			return mip;
		
		}
		else if(PortType == "DataInPort")
		{
			T *m_in = new T();
			InPort<T> *m_inIn = new InPort<T>(tname.c_str(),*m_in);

			addInPort(tname.c_str(), *m_inIn);

			MyPortBase *mip = new MyInPortSeq<T,T2>(m_in,m_inIn,tname,tdt,this);
			
			

			mip->SetExcelParam(c,l,sn,leng,mstate);
			ConfInPorts.push_back(mip);
			

			

			return mip;
		}
		return NULL;
	}
	//�R���t�B�M�����[�V�����p�����[�^�ɂ��f�[�^�|�[�g���쐬����֐�
   template <typename T, typename T2> MyPortBase* ConfcrPort(std::string tname, std::string PortType, std::string tdt, int c, std::string l, std::string sn, std::string leng, bool mstate)
   {
	  
	  
		if(PortType == "DataOutPort")
		{
			T *m_out = new T();
			OutPort<T> *m_outOut = new OutPort<T>(tname.c_str(),*m_out);
			addOutPort(tname.c_str(), *m_outOut);
			
			MyPortBase *mip;
			
			mip= new MyOutPort<T,T2>(m_out,m_outOut,tname,tdt,this);
			
			mip->SetExcelParam(c,l,sn,leng,mstate);
			ConfOutPorts.push_back(mip);
			

			

			return mip;
		
		}
		else if(PortType == "DataInPort")
		{
			T *m_in = new T();
			InPort<T> *m_inIn = new InPort<T>(tname.c_str(),*m_in);
			

			addInPort(tname.c_str(), *m_inIn);

			MyPortBase *mip = new MyInPort<T,T2>(m_in,m_inIn,tname,tdt,this);
			
			

			mip->SetExcelParam(c,l,sn,leng,mstate);
			ConfInPorts.push_back(mip);
			

			

			return mip;
		}
		return NULL;
   }


   

   
   
   //�f�[�^�|�[�g�̍쐬�A�l�̐ݒ���s���֐�
   MyPortBase* SetDPort(std::vector<std::string> pt, int c, std::string l, std::string sn, std::string leng, bool mstate, bool msflag);
   //�f�[�^�|�[�g���폜����֐�
   void DelDPort(std::vector<std::string> pt);
   //�f�[�^�|�[�g���擾����֐�
   MyPortBase* GetDPort(std::vector<std::string> pt);
   
   //�f�[�^�|�[�g���֘A�t����֐�
   void AttachPort(MyPortBase *mpb, std::string n);
   //�f�[�^�|�[�g�̊֘A�t������������֐�
   void DetachPort(MyPortBase *mpb, std::string n);
   //�C���|�[�g���擾����֐�
   MyPortBase *GetInPort(std::string n);
   //�A�E�g�|�[�g���擾����֐�
   MyPortBase *GetOutPort(std::string n);
   //�R���t�B�M�����[�V�����p�����[�^�Őݒ肵���f�[�^�|�[�g���擾����֐�
   MyPortBase *GetConfOutPort(std::string n);
   MyPortBase *GetConfInPort(std::string n);

   //RTC�̃f�[�^�|�[�g�̃c���[���擾����֐�
   TreeObject* GetRTCTree(std::string IP_adress);

   std::vector<MyPortBase*> InPorts;
   std::vector<MyPortBase*> OutPorts;

   std::vector<MyPortBase*> ConfInPorts;
   std::vector<MyPortBase*> ConfOutPorts;

   RTC::Manager* m_manager;

   std::vector<OtherPort> rtclist;

   RTC::CorbaPort m_DataBasePort;
   mDataBaseSVC_impl m_database;

   
   

 protected:
	std::string file_path;
	std::string conf_data_type;
	std::string conf_port_type;
	int conf_column;
	std::string conf_start_row;
	std::string conf_end_row;
	std::string conf_sheetname;


 private:


};


extern "C"
{
  DLL_EXPORT void ExcelRTCInit(RTC::Manager* manager);
};

#endif