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
*@brief インポートがデータを受けた時のコールバック
*/
template <class T, class T2>
class DataListener
  : public ConnectorDataListenerT<T>
{
public:
	/**
	*@brief コンストラクタ
	* @param mp データポートオブジェクト
	* @param sv データ格納用コンテナ
	* @param mrtc
	*/
	DataListener(MyPortBase *mp, std::vector<std::vector<T2>> &sv, ExcelRTComponent *mrtc){
		m_port = mp;
		m_data = &sv;
		m_rtc = mrtc;
	}
	/**
	*@brief デストラクタ
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
  MyPortBase *m_port;	/**<　@brief  */
  std::vector<std::vector<T2>> *m_data; /**<　@brief  */
  ExcelRTComponent *m_rtc; /**<　@brief  */
  
};


/**
 * @class SeqDataListener
*@brief シーケンス型のインポートがデータを受け取ったときのコールバック
*/
template <class T, class T2>
class SeqDataListener
  : public ConnectorDataListenerT<T>
{
public:
	/**
	*@brief コンストラクタ
	* @param mp データポートオブジェクト
	* @param sv データ格納用コンテナ
	* @param mrtc
	*/
	SeqDataListener(MyPortBase *mp, std::vector<std::vector<T2>> &sv, ExcelRTComponent *mrtc){
		m_port = mp;
		m_data = &sv;
		m_rtc = mrtc;
	}
	/**
	*@brief デストラクタ
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
  MyPortBase *m_port; /**<　@brief  */
  std::vector<std::vector<T2>> *m_data; /**<　@brief  */
  ExcelRTComponent *m_rtc; /**<　@brief  */
  
};



/**
 * @class OtherPort
*@brief 他のRTCのデータポートのクラス
*/
class OtherPort
{
	public:
		/**
		*@brief コンストラクタ
		* @param p
		* @param s
		*/
		OtherPort(PortService_var p, std::vector<std::string> s)
		{
			pb = p;
			buff = s;
			mpb = NULL;
		};
		PortService_var pb; /**<　@brief  */
		std::vector<std::string> buff; /**<　@brief  */
		MyPortBase *mpb; /**<　@brief  */
};



/**
 * @class MyPortBase
*@brief ExcelRTCのデータポートのクラス
*/
class MyPortBase
{
public:
	/**
	*@brief デストラクタ
	*/
	virtual ~MyPortBase()
	{
		delete pb;
	}

	
	/**
	*@brief GUIで入力した値を設定する関数
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
	*@brief データを取得したかを判定する関数
	*/
	virtual bool isNew()
	{
		return false;
	}
	/**
	*@brief 受けたデータをセルに書き込む関数
	* @param moption
	*/
	virtual void PutData(bool moption)
	{

	}
	/**
	*@brief データポートと関連付けしたセルに名前を入力
	*/
	virtual void update_cellName()
	{
		
		myExcel::Obj->SetCellStringSingle(col, low, sheetName, name);
	}


	std::string name; /**<　@brief  */
	
	int col; /**<　@brief  */
	std::string low; /**<　@brief  */
	std::string length; /**<　@brief  */
	int sheetnum; /**<　@brief  */
	
	std::string sheetName; /**<　@brief  */
	int num; /**<　@brief  */
	OtherPort *mop; /**<　@brief  */
	PortBase *pb; /**<　@brief  */
	std::string data_type; /**<　@brief  */
	std::vector<std::string> attachPort;  /**<　@brief  */

	bool state; /**<　@brief  */
	ExcelRTComponent *mexc; /**<　@brief  */
	coil::Mutex _mutex; /**<　@brief  */

	
	

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
*@brief インポートのクラス
*/
template <class T, class T2>
class MyInPort : public MyPortBase
{
public:
	/**
	*@brief コンストラクタ
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
	*@brief デストラクタ
	*/
	~MyInPort()
	{
		//delete inIn;
		//delete In;
	}
	
	/**
	*@brief データを取得したかを判定する関数
	*/
	virtual bool isNew()
	{
		return inIn->isNew();
	}

	
	
	/**
	*@brief 受けたデータをセルに書き込む関数
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

	T *In; /**<　@brief  */
	RTC::InPort<T> *inIn; /**<　@brief  */
	std::vector<std::vector<T2>> buff; /**<　@brief  */
	
};


/**
 * @class MyInPortSeq
*@brief シーケンス型のインポートのクラス
*/
template <class T, class T2>
class MyInPortSeq : public MyPortBase
{
public:
	/**
	*@brief コンストラクタ
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
	*@brief デストラクタ
	*/
	~MyInPortSeq()
	{
		//delete inIn;
		//delete In;
	}
	
	/**
	*@brief データを取得したかを判定する関数
	*/
	virtual bool isNew()
	{
		return inIn->isNew();
	}

	
		

		
		
	
	/**
	*@brief 受けたデータをセルに書き込む関数
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

	T *In; /**<　@brief  */
	RTC::InPort<T> *inIn; /**<　@brief  */
	std::vector<std::vector<T2>> buff; /**<　@brief  */
	
};



/**
 * @class MyOutPort
*@brief アウトポートのクラス
*/
template <class T, class T2>
class MyOutPort : public MyPortBase
{
public:
	/**
	*@brief コンストラクタ
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
	*@brief デストラクタ
	*/
	~MyOutPort()
	{
		//delete outOut;
		//delete Out;
	}
	
	/**
	*@brief セルの値をデータポートから出力する関数
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
	T *Out; /**<　@brief  */
	RTC::OutPort<T> *outOut; /**<　@brief  */
};


/**
 * @class MyOutPortSeq
*@brief シーケンス型のアウトポートのクラス
*/
template <class T, class T2>
class MyOutPortSeq : public MyPortBase
{
public:
	/**
	*@brief コンストラクタ
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
	*@brief デストラクタ
	*/
	~MyOutPortSeq()
	{
		//delete outOut;
		//delete Out;
	}


	
	/**
	*@brief セルの値をデータポートから出力する関数
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
	T *Out; /**<　@brief  */
	RTC::OutPort<T> *outOut; /**<　@brief  */
};


/**
*@brief データポートを接続
* @param p1
* @param p2
*/
void portConnect(PortService_ptr p1, PortService_var p2);



/**
 * @class ExcelRTComponent
*@brief Excelを操作するRTコンポーネント
*/
class ExcelRTComponent
  : public RTC::DataFlowComponentBase
{
 public:
  
	/**
	*@brief コンストラクタ
	* @param manager
	*/
  ExcelRTComponent(RTC::Manager* manager);

  /**
	*@brief デストラクタ
	*/
  ~ExcelRTComponent();

   
  /**
  *@brief 初期化処理用コールバック関数
  * @return 
  */
   virtual RTC::ReturnCode_t onInitialize();

   
   /**
   *@brief 周期処理用コールバック関数
   * @param ec_id
   * @return 
   */
   virtual RTC::ReturnCode_t onExecute(RTC::UniqueId ec_id);
   
   
   /**
   *@brief 不活性化時のコールバック関数
   * @param ec_id
   * @return 
   */
   virtual RTC::ReturnCode_t onDeactivated(RTC::UniqueId ec_id);

   /**
   *@brief 終了処理のコールバック関数
   * @return 
   */
   virtual RTC::ReturnCode_t onFinalize();

   
   /**
   *@brief データポートを作成する関数
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
   *@brief データポートを削除する関数
   * @param op
   */
   void DeleteOtherPort(OtherPort &op);
   /**
   *@brief 全てのデータポートを削除する関数
   */
   void DeleteAllPort();
   
   /**
   *@brief RTCの情報を保存用シートに書き込む関数
   */
   void Save();
   
   /**
   *@brief RTCの情報をほ場用シートより読み込む関数
   */
   void Load();
   
   /**
   *@brief データを書き込む列を初期化する
   * @param mpb
   */
   void ResetPort(MyPortBase* mpb);
   
   /**
   *@brief 全てのデータを書き込む列を初期化する
   */
   void ResetAllPort();
   
   /**
   *@brief コンフィギュレーションパラメータが変更されたときに呼び出される関数
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
   *@brief セルに値を書き込む関数
   * @param dt
   * @param pb
   */
    template <typename T>
   void SetCellData(std::vector<std::vector<T>> dt, MyPortBase *pb)
	{
		
		myExcel::Obj->SetCellValue<T>(pb->col+pb->num, pb->low, pb->sheetName, pb->state, dt);
		
	}
   
   /**
   *@brief セルの値を取得する関数
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
	*@brief シーケンス型のデータポートを作成する関数
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
	*@brief データポートを作成する関数
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
   *@brief コンフィギュレーションパラメータによりシーケンス型のデータポートを作成する関数
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
   *@brief コンフィギュレーションパラメータによりデータポートを作成する関数
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
   *@brief データポートの作成、値の設定を行う関数
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
   *@brief データポートを削除する関数
   * @param pt
   */
   void DelDPort(std::vector<std::string> pt);
   
   /**
   *@brief データポートを取得する関数
   * @param pt
   * @return 
   */
   MyPortBase* GetDPort(std::vector<std::string> pt);
   
   
   /**
   *@brief データポートを関連付ける関数
   * @param mpb
   * @param n
   */
   void AttachPort(MyPortBase *mpb, std::string n);
   
   /**
   *@brief データポートの関連付けを解除する関数
   * @param mpb
   * @param n
   */
   void DetachPort(MyPortBase *mpb, std::string n);
   
   /**
   *@brief インポートを取得する関数
   * @param n
   * @return 
   */
   MyPortBase *GetInPort(std::string n);
   
   /**
   *@brief アウトポートを取得する関数
   * @param n
   * @return 
   */
   MyPortBase *GetOutPort(std::string n);
   
   /**
   *@brief コンフィギュレーションパラメータで設定したアウトポートを取得する関数
   * @param n
   * @return 
   */
   MyPortBase *GetConfOutPort(std::string n);
   /**
   *@brief コンフィギュレーションパラメータで設定したインポートを取得する関数
   * @param n
   * @return 
   */
   MyPortBase *GetConfInPort(std::string n);

   
   /**
   *@brief RTCのデータポートのツリーを取得する関数
   * @param IP_adress
   * @return 
   */
   TreeObject* GetRTCTree(std::string IP_adress);

   std::vector<MyPortBase*> InPorts;	/**<　@brief  */
   std::vector<MyPortBase*> OutPorts;	/**<　@brief  */

   std::vector<MyPortBase*> ConfInPorts;	/**<　@brief  */
   std::vector<MyPortBase*> ConfOutPorts;	/**<　@brief  */

   RTC::Manager* m_manager;	/**<　@brief  */

   std::vector<OtherPort> rtclist;	/**<　@brief  */

   RTC::CorbaPort m_SpreadSheetPort;	/**<　@brief  */
   mSpreadSheetSVC_impl m_spreadsheet;	/**<　@brief  */

   
   

 protected:
	std::string file_path;	/**<　@brief  */
	std::string conf_data_type;	/**<　@brief  */
	std::string conf_port_type;	/**<　@brief  */
	int conf_column;	/**<　@brief  */
	std::string conf_start_row;	/**<　@brief  */
	std::string conf_end_row;	/**<　@brief  */
	std::string conf_sheetname;	/**<　@brief  */

	int actionLock;	/**<　@brief  */
	int Red;	/**<　@brief  */
	int Green;	/**<　@brief  */
	int Blue;	/**<　@brief  */

	coil::Mutex _mutex; /**<　@brief  */


 private:


};


extern "C"
{
  DLL_EXPORT void ExcelRTCInit(RTC::Manager* manager);
};

#endif