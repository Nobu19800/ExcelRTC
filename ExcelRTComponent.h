// -*-C++-*-
/*!
 * @file  ExcelRTComponent.h
 *
 */
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

#include "SpreadSheetSVC_impl.h"



class ExcelRTComponent;
class MyPortBase;
class OtherPort;
class MyOnUpdate;



void ListRecursive(CosNaming::NamingContext_ptr context,std::vector<OtherPort> &rtclist,std::string &name, TreeObject *to);

int rtc_get_rtclist(RTC::CorbaNaming &naming,std::vector<OtherPort> &rtclist, TreeObject *to, std::string IP_adress);






/**
 * @class DataListener
*@brief �C���|�[�g���f�[�^���󂯂����̃R�[���o�b�N
*/
template <class T, class T2>
class DataListener
  : public ConnectorDataListenerT<T>
{
public:
	/**
	*@brief �R���X�g���N�^
	* @param mp �f�[�^�|�[�g�I�u�W�F�N�g
	* @param sv �f�[�^�i�[�p�R���e�i
	* @param mrtc
	*/
	DataListener(MyPortBase *mp, std::vector<std::vector<T2>> &sv, ExcelRTComponent *mrtc){
		m_port = mp;
		m_data = &sv;
		m_rtc = mrtc;
	}
	/**
	*@brief �f�X�g���N�^
	*/
  virtual ~DataListener()
  {
  
  }

  /**
	*@brief
	* @param info
	* @param data
	*/
  virtual void operator()(const ConnectorInfo& info,
                          const T& data)
  {

	 

	  std::vector<T2> tmp;
	
	  
	  tmp.push_back(data.data);
	
	  m_port->_mutex.lock();
	  m_data->push_back(tmp);
	  m_port->_mutex.unlock();

	  m_rtc->updateAPort(m_port);
       
  };
  MyPortBase *m_port;	/**<�@@brief  */
  std::vector<std::vector<T2>> *m_data; /**<�@@brief  */
  ExcelRTComponent *m_rtc; /**<�@@brief  */
  
};


/**
 * @class SeqDataListener
*@brief �V�[�P���X�^�̃C���|�[�g���f�[�^���󂯎�����Ƃ��̃R�[���o�b�N
*/
template <class T, class T2>
class SeqDataListener
  : public ConnectorDataListenerT<T>
{
public:
	/**
	*@brief �R���X�g���N�^
	* @param mp �f�[�^�|�[�g�I�u�W�F�N�g
	* @param sv �f�[�^�i�[�p�R���e�i
	* @param mrtc
	*/
	SeqDataListener(MyPortBase *mp, std::vector<std::vector<T2>> &sv, ExcelRTComponent *mrtc){
		m_port = mp;
		m_data = &sv;
		m_rtc = mrtc;
	}
	/**
	*@brief �f�X�g���N�^
	*/
  virtual ~SeqDataListener()
  {
  
  }

  /**
	*@brief
	* @param info
	* @param data
	*/
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

	m_rtc->updateAPort(m_port);
  };
  MyPortBase *m_port; /**<�@@brief  */
  std::vector<std::vector<T2>> *m_data; /**<�@@brief  */
  ExcelRTComponent *m_rtc; /**<�@@brief  */
  
};



/**
 * @class OtherPort
*@brief ����RTC�̃f�[�^�|�[�g�̃N���X
*/
class OtherPort
{
	public:
		/**
		*@brief �R���X�g���N�^
		* @param p
		* @param s
		*/
		OtherPort(PortService_var p, std::vector<std::string> s)
		{
			pb = p;
			buff = s;
			mpb = NULL;
		};
		PortService_var pb; /**<�@@brief  */
		std::vector<std::string> buff; /**<�@@brief  */
		MyPortBase *mpb; /**<�@@brief  */
};



/**
 * @class MyPortBase
*@brief ExcelRTC�̃f�[�^�|�[�g�̃N���X
*/
class MyPortBase
{
public:
	/**
	*@brief �f�X�g���N�^
	*/
	virtual ~MyPortBase()
	{
		delete pb;
	}

	
	/**
	*@brief GUI�œ��͂����l��ݒ肷��֐�
	* @param c
	* @param l
	* @param sn
	* @param leng
	* @param mstate
	*/
	virtual void SetExcelParam(int c, std::string l, std::string sn, std::string leng, bool mstate)
	{
		col = c;
		low = l;
		sheetName = sn;
		length = leng;
		num = 0;
		state = mstate;
		update_cellName();
	}

	/**
	*@brief �f�[�^���擾�������𔻒肷��֐�
	*/
	virtual bool isNew()
	{
		return false;
	}
	/**
	*@brief �󂯂��f�[�^���Z���ɏ������ފ֐�
	* @param moption
	*/
	virtual void PutData(bool moption)
	{

	}
	/**
	*@brief �f�[�^�|�[�g�Ɗ֘A�t�������Z���ɖ��O�����
	*/
	virtual void update_cellName()
	{
		
		myExcel::Obj->SetCellStringSingle(col, low, sheetName, name);
	}


	std::string name; /**<�@@brief  */
	
	int col; /**<�@@brief  */
	std::string low; /**<�@@brief  */
	std::string length; /**<�@@brief  */
	int sheetnum; /**<�@@brief  */
	
	std::string sheetName; /**<�@@brief  */
	int num; /**<�@@brief  */
	OtherPort *mop; /**<�@@brief  */
	PortBase *pb; /**<�@@brief  */
	std::string data_type; /**<�@@brief  */
	std::vector<std::string> attachPort;  /**<�@@brief  */

	bool state; /**<�@@brief  */
	ExcelRTComponent *mexc; /**<�@@brief  */
	coil::Mutex _mutex; /**<�@@brief  */

	
	

	/*template <typename T>T GetOutPort()
	{
		return static_cast<MyOutPort<T>*>this;
	}

	template <typename T>T GetInPort()
	{
		return static_cast<MyInPort<T>*>this;
	}*/
};


/**
 * @class MyInPort
*@brief �C���|�[�g�̃N���X
*/
template <class T, class T2>
class MyInPort : public MyPortBase
{
public:
	/**
	*@brief �R���X�g���N�^
	* @param id
	* @param ip
	* @param n
	* @param dt
	* @param m_mexc
	*/
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

		
		inIn->addConnectorDataListener(ON_BUFFER_WRITE, new DataListener<T,T2>(this, buff, mexc));
		
	}
	/**
	*@brief �f�X�g���N�^
	*/
	~MyInPort()
	{
		//delete inIn;
		//delete In;
	}
	
	/**
	*@brief �f�[�^���擾�������𔻒肷��֐�
	*/
	virtual bool isNew()
	{
		return inIn->isNew();
	}

	
	
	/**
	*@brief �󂯂��f�[�^���Z���ɏ������ފ֐�
	* @param moption
	*/
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

	T *In; /**<�@@brief  */
	RTC::InPort<T> *inIn; /**<�@@brief  */
	std::vector<std::vector<T2>> buff; /**<�@@brief  */
	
};


/**
 * @class MyInPortSeq
*@brief �V�[�P���X�^�̃C���|�[�g�̃N���X
*/
template <class T, class T2>
class MyInPortSeq : public MyPortBase
{
public:
	/**
	*@brief �R���X�g���N�^
	* @param id
	* @param ip
	* @param n
	* @param dt
	* @param m_mexc
	*/
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

		
		inIn->addConnectorDataListener(ON_BUFFER_WRITE, new SeqDataListener<T,T2>(this, buff, mexc));


	}
	/**
	*@brief �f�X�g���N�^
	*/
	~MyInPortSeq()
	{
		//delete inIn;
		//delete In;
	}
	
	/**
	*@brief �f�[�^���擾�������𔻒肷��֐�
	*/
	virtual bool isNew()
	{
		return inIn->isNew();
	}

	
		

		
		
	
	/**
	*@brief �󂯂��f�[�^���Z���ɏ������ފ֐�
	* @param moption
	*/
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

	T *In; /**<�@@brief  */
	RTC::InPort<T> *inIn; /**<�@@brief  */
	std::vector<std::vector<T2>> buff; /**<�@@brief  */
	
};



/**
 * @class MyOutPort
*@brief �A�E�g�|�[�g�̃N���X
*/
template <class T, class T2>
class MyOutPort : public MyPortBase
{
public:
	/**
	*@brief �R���X�g���N�^
	* @param od
	* @param op
	* @param n
	* @param dt
	* @param m_mexc
	*/
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
	/**
	*@brief �f�X�g���N�^
	*/
	~MyOutPort()
	{
		//delete outOut;
		//delete Out;
	}
	
	/**
	*@brief �Z���̒l���f�[�^�|�[�g����o�͂���֐�
	* @param moption
	*/
	virtual void PutData(bool moption)
	{
				
		std::vector<T2> v = mexc->GetCellData<T2>(this);
		
		Out->data = v[0];
		if(state)
			num += 1;
		
		
		outOut->write();
	}
	T *Out; /**<�@@brief  */
	RTC::OutPort<T> *outOut; /**<�@@brief  */
};


/**
 * @class MyOutPortSeq
*@brief �V�[�P���X�^�̃A�E�g�|�[�g�̃N���X
*/
template <class T, class T2>
class MyOutPortSeq : public MyPortBase
{
public:
	/**
	*@brief �R���X�g���N�^
	* @param od
	* @param op
	* @param n
	* @param dt
	* @param m_mexc
	*/
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
	/**
	*@brief �f�X�g���N�^
	*/
	~MyOutPortSeq()
	{
		//delete outOut;
		//delete Out;
	}


	
	/**
	*@brief �Z���̒l���f�[�^�|�[�g����o�͂���֐�
	* @param moption
	*/
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
	T *Out; /**<�@@brief  */
	RTC::OutPort<T> *outOut; /**<�@@brief  */
};


/**
*@brief �f�[�^�|�[�g��ڑ�
* @param p1
* @param p2
*/
void portConnect(PortService_ptr p1, PortService_var p2);



/**
 * @class ExcelRTComponent
*@brief Excel�𑀍삷��RT�R���|�[�l���g
*/
class ExcelRTComponent
  : public RTC::DataFlowComponentBase
{
 public:
  
	/**
	*@brief �R���X�g���N�^
	* @param manager
	*/
  ExcelRTComponent(RTC::Manager* manager);

  /**
	*@brief �f�X�g���N�^
	*/
  ~ExcelRTComponent();

   
  /**
  *@brief �����������p�R�[���o�b�N�֐�
  * @return 
  */
   virtual RTC::ReturnCode_t onInitialize();

   
   /**
   *@brief ���������p�R�[���o�b�N�֐�
   * @param ec_id
   * @return 
   */
   virtual RTC::ReturnCode_t onExecute(RTC::UniqueId ec_id);
   
   
   /**
   *@brief �s���������̃R�[���o�b�N�֐�
   * @param ec_id
   * @return 
   */
   virtual RTC::ReturnCode_t onDeactivated(RTC::UniqueId ec_id);

   /**
   *@brief �I�������̃R�[���o�b�N�֐�
   * @return 
   */
   virtual RTC::ReturnCode_t onFinalize();

   
   /**
   *@brief �f�[�^�|�[�g���쐬����֐�
   * @param op
   * @param c
   * @param l
   * @param sn
   * @param leng
   * @param mstate
   * @return 
   */
   MyPortBase* CreatePort(OtherPort &op, int c, std::string l, std::string sn, std::string leng, bool mstate);
   
   /**
   *@brief �f�[�^�|�[�g���폜����֐�
   * @param op
   */
   void DeleteOtherPort(OtherPort &op);
   /**
   *@brief �S�Ẵf�[�^�|�[�g���폜����֐�
   */
   void DeleteAllPort();
   
   /**
   *@brief RTC�̏���ۑ��p�V�[�g�ɏ������ފ֐�
   */
   void Save();
   
   /**
   *@brief RTC�̏����ُ�p�V�[�g���ǂݍ��ފ֐�
   */
   void Load();
   
   /**
   *@brief �f�[�^���������ޗ������������
   * @param mpb
   */
   void ResetPort(MyPortBase* mpb);
   
   /**
   *@brief �S�Ẵf�[�^���������ޗ������������
   */
   void ResetAllPort();
   
   /**
   *@brief �R���t�B�M�����[�V�����p�����[�^���ύX���ꂽ�Ƃ��ɌĂяo�����֐�
   */
   void ConfigUpdate();

   /**
   *@brief
   */
   void update_cellName();

   /**
   *@brief
   */
   void updateAPort(MyPortBase* ip);

   
   /**
   *@brief �Z���ɒl���������ފ֐�
   * @param dt
   * @param pb
   */
    template <typename T>
   void SetCellData(std::vector<std::vector<T>> dt, MyPortBase *pb)
	{
		
		myExcel::Obj->SetCellValue<T>(pb->col+pb->num, pb->low, pb->sheetName, pb->state, dt);
		
	}
   
   /**
   *@brief �Z���̒l���擾����֐�
   * @param pb
   * @return 
   */
    template <typename T>
	std::vector<T> GetCellData(MyPortBase *pb)
	{
		std::vector<T> td = myExcel::Obj->GetCellValue<T>(pb->col+pb->num, pb->low, pb->sheetName, pb->length);
		
		return td;
	}
	
	/**
	*@brief �V�[�P���X�^�̃f�[�^�|�[�g���쐬����֐�
	* @param op
	* @param tdt
	* @param c
	* @param l
	* @param sn
	* @param leng
	* @param mstate
	* @return 
	*/
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
	
	/**
	*@brief �f�[�^�|�[�g���쐬����֐�
	* @param op
	* @param tdt
	* @param c
	* @param l
	* @param sn
	* @param leng
	* @param mstate
	* @return 
	*/
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
   
   /**
   *@brief �R���t�B�M�����[�V�����p�����[�^�ɂ��V�[�P���X�^�̃f�[�^�|�[�g���쐬����֐�
    * @param tname
	* @param PortType
	* @param tdt
	* @param c
	* @param l
	* @param sn
	* @param leng
	* @param mstate
	* @return 
   */
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
	
   /**
   *@brief �R���t�B�M�����[�V�����p�����[�^�ɂ��f�[�^�|�[�g���쐬����֐�
    * @param tname
	* @param PortType
	* @param tdt
	* @param c
	* @param l
	* @param sn
	* @param leng
	* @param mstate
    * @return 
   */
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


   

   
   
   
   /**
   *@brief �f�[�^�|�[�g�̍쐬�A�l�̐ݒ���s���֐�
    * @param pt
	* @param c
	* @param l
	* @param sn
	* @param leng
	* @param mstate
	* @param msflag
	* @return 
   */
   MyPortBase* SetDPort(std::vector<std::string> pt, int c, std::string l, std::string sn, std::string leng, bool mstate, bool msflag);
   
   /**
   *@brief �f�[�^�|�[�g���폜����֐�
   * @param pt
   */
   void DelDPort(std::vector<std::string> pt);
   
   /**
   *@brief �f�[�^�|�[�g���擾����֐�
   * @param pt
   * @return 
   */
   MyPortBase* GetDPort(std::vector<std::string> pt);
   
   
   /**
   *@brief �f�[�^�|�[�g���֘A�t����֐�
   * @param mpb
   * @param n
   */
   void AttachPort(MyPortBase *mpb, std::string n);
   
   /**
   *@brief �f�[�^�|�[�g�̊֘A�t������������֐�
   * @param mpb
   * @param n
   */
   void DetachPort(MyPortBase *mpb, std::string n);
   
   /**
   *@brief �C���|�[�g���擾����֐�
   * @param n
   * @return 
   */
   MyPortBase *GetInPort(std::string n);
   
   /**
   *@brief �A�E�g�|�[�g���擾����֐�
   * @param n
   * @return 
   */
   MyPortBase *GetOutPort(std::string n);
   
   /**
   *@brief �R���t�B�M�����[�V�����p�����[�^�Őݒ肵���A�E�g�|�[�g���擾����֐�
   * @param n
   * @return 
   */
   MyPortBase *GetConfOutPort(std::string n);
   /**
   *@brief �R���t�B�M�����[�V�����p�����[�^�Őݒ肵���C���|�[�g���擾����֐�
   * @param n
   * @return 
   */
   MyPortBase *GetConfInPort(std::string n);

   
   /**
   *@brief RTC�̃f�[�^�|�[�g�̃c���[���擾����֐�
   * @param IP_adress
   * @return 
   */
   TreeObject* GetRTCTree(std::string IP_adress);

   std::vector<MyPortBase*> InPorts;	/**<�@@brief  */
   std::vector<MyPortBase*> OutPorts;	/**<�@@brief  */

   std::vector<MyPortBase*> ConfInPorts;	/**<�@@brief  */
   std::vector<MyPortBase*> ConfOutPorts;	/**<�@@brief  */

   RTC::Manager* m_manager;	/**<�@@brief  */

   std::vector<OtherPort> rtclist;	/**<�@@brief  */

   RTC::CorbaPort m_SpreadSheetPort;	/**<�@@brief  */
   mSpreadSheetSVC_impl m_spreadsheet;	/**<�@@brief  */

   
   

 protected:
	std::string file_path;	/**<�@@brief  */
	std::string conf_data_type;	/**<�@@brief  */
	std::string conf_port_type;	/**<�@@brief  */
	int conf_column;	/**<�@@brief  */
	std::string conf_start_row;	/**<�@@brief  */
	std::string conf_end_row;	/**<�@@brief  */
	std::string conf_sheetname;	/**<�@@brief  */

	int actionLock;	/**<�@@brief  */
	int Red;	/**<�@@brief  */
	int Green;	/**<�@@brief  */
	int Blue;	/**<�@@brief  */

	coil::Mutex _mutex; /**<�@@brief  */


 private:


};


extern "C"
{
  DLL_EXPORT void ExcelRTCInit(RTC::Manager* manager);
};

#endif