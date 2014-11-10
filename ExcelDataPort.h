// -*-C++-*-
/*!
 * @file  ExcelDataPort.h
 * @brief ExcelDataPort
 */
#ifndef EXCELDATAPORT_H
#define EXCELDATAPORT_H


#include "ExcelControl.h"

class ExcelPortBase;
class ExcelControl;


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
	* @param mrtc ExcelRTC
	*/
	DataListener(ExcelPortBase *mp, std::vector<std::vector<T2>> &sv, ExcelControl *mrtc){
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
	* @param info コネクタの情報
	* @param data データ
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
  ExcelPortBase *m_port;	/**<　@brief  */
  std::vector<std::vector<T2>> *m_data; /**<　@brief  */
  ExcelControl *m_rtc; /**<　@brief  */
  
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
	* @param mrtc ExcelRTC
	*/
	SeqDataListener(ExcelPortBase *mp, std::vector<std::vector<T2>> &sv, ExcelControl *mrtc){
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
	* @param info コネクタの情報
	* @param data データ
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
  ExcelPortBase *m_port; /**<　@brief  */
  std::vector<std::vector<T2>> *m_data; /**<　@brief  */
  ExcelControl *m_rtc; /**<　@brief  */
  
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
		* @param p データポートオブジェクト
		* @param s パス
		*/
		OtherPort(PortService_var p, std::vector<std::string> s)
		{
			pb = p;
			buff = s;
			mpb = NULL;
		};
		PortService_var pb; /**<　@brief  */
		std::vector<std::string> buff; /**<　@brief  */
		ExcelPortBase *mpb; /**<　@brief  */
};



/**
 * @class ExcelPortBase
*@brief ExcelRTCのデータポートのクラス
*/
class ExcelPortBase
{
public:
	/**
	*@brief デストラクタ
	*/
	virtual ~ExcelPortBase()
	{
		delete pb;
	}

	
	/**
	*@brief GUIで入力した値を設定する関数
	* @param c 列番号
	* @param l 行番号
	* @param sn シート名
	* @param leng 行の範囲
	* @param mstate 列を移動する場合はTrue
	*/
	virtual void setParam(int c, std::string l, std::string sn, std::string leng, bool mstate)
	{
		col = c;
		low = l;
		sheetName = sn;
		length = leng;
		num = 0;
		state = mstate;
		//update_cellName();
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
	* @param moption Trueの場合はバッファのデータをまとめて処理する。Falseの場合は1回だけデータを読み込む。
	*/
	virtual void putData(bool moption)
	{

	}
	/**
	*@brief データポートと関連付けしたセルに名前を入力
	*/
	virtual void update_cellName()
	{
		
		ExcelObject::Obj->setCellStringSingle(col, low, sheetName, name);
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
	ExcelControl *mexc; /**<　@brief  */
	coil::Mutex _mutex; /**<　@brief  */

	
	

	/*template <typename T>T GetOutPort()
	{
		return static_cast<ExcelOutPort<T>*>this;
	}

	template <typename T>T GetInPort()
	{
		return static_cast<ExcelInPort<T>*>this;
	}*/
};


/**
 * @class ExcelInPort
*@brief インポートのクラス
*/
template <class T, class T2>
class ExcelInPort : public ExcelPortBase
{
public:
	/**
	*@brief コンストラクタ
	* @param id データオブジェクト
	* @param ip データポートオブジェクト
	* @param n 名称
	* @param dt データ型
	* @param m_mexc ExcelRTC
	*/
	ExcelInPort(T *id, RTC::InPort<T> *ip, std::string n, std::string dt, ExcelControl *m_mexc){
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
	~ExcelInPort()
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
	* @param moption Trueの場合はバッファのデータをまとめて処理する。Falseの場合は1回だけデータを読み込む。
	*/
	virtual void putData(bool moption)
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
		
		

		mexc->setCellData<T2>(v, this);

		if(state)
			num += v.size();
		

		

		
	}

	T *In; /**<　@brief  */
	RTC::InPort<T> *inIn; /**<　@brief  */
	std::vector<std::vector<T2>> buff; /**<　@brief  */
	
};


/**
 * @class ExcelInPortSeq
*@brief シーケンス型のインポートのクラス
*/
template <class T, class T2>
class ExcelInPortSeq : public ExcelPortBase
{
public:
	/**
	*@brief コンストラクタ
	* @param id データオブジェクト
	* @param ip データポートオブジェクト
	* @param n 名称
	* @param dt データ型
	* @param m_mexc ExcelRTC
	*/
	ExcelInPortSeq(T *id, RTC::InPort<T> *ip, std::string n, std::string dt, ExcelControl *m_mexc){
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
	~ExcelInPortSeq()
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
	* @param moption Trueの場合はバッファのデータをまとめて処理する。Falseの場合は1回だけデータを読み込む。
	*/
	virtual void putData(bool moption)
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

		mexc->setCellData<T2>(v, this);

		if(state)
			num += v.size();

		
				
				

			
		
	}

	T *In; /**<　@brief  */
	RTC::InPort<T> *inIn; /**<　@brief  */
	std::vector<std::vector<T2>> buff; /**<　@brief  */
	
};



/**
 * @class ExcelOutPort
*@brief アウトポートのクラス
*/
template <class T, class T2>
class ExcelOutPort : public ExcelPortBase
{
public:
	/**
	*@brief コンストラクタ
	* @param id データオブジェクト
	* @param ip データポートオブジェクト
	* @param n 名称
	* @param dt データ型
	* @param m_mexc ExcelRTC
	*/
	ExcelOutPort(T *od, RTC::OutPort<T> *op, std::string n, std::string dt, ExcelControl *m_mexc){
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
	~ExcelOutPort()
	{
		//delete outOut;
		//delete Out;
	}
	
	/**
	*@brief セルの値をデータポートから出力する関数
	* @param moption Trueの場合はバッファのデータをまとめて処理する。Falseの場合は1回だけデータを読み込む。
	*/
	virtual void putData(bool moption)
	{
				
		std::vector<T2> v = mexc->getCellData<T2>(this);
		
		Out->data = v[0];
		if(state)
			num += 1;
		
		
		outOut->write();
	}
	T *Out; /**<　@brief  */
	RTC::OutPort<T> *outOut; /**<　@brief  */
};


/**
 * @class ExcelOutPortSeq
*@brief シーケンス型のアウトポートのクラス
*/
template <class T, class T2>
class ExcelOutPortSeq : public ExcelPortBase
{
public:
	/**
	*@brief コンストラクタ
	* @param id データオブジェクト
	* @param ip データポートオブジェクト
	* @param n 名称
	* @param dt データ型
	* @param m_mexc ExcelRTC
	*/
	ExcelOutPortSeq(T *od, RTC::OutPort<T> *op, std::string n, std::string dt, ExcelControl *m_mexc){
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
	~ExcelOutPortSeq()
	{
		//delete outOut;
		//delete Out;
	}


	
	/**
	*@brief セルの値をデータポートから出力する関数
	* @param moption Trueの場合はバッファのデータをまとめて処理する。Falseの場合は1回だけデータを読み込む。
	*/
	virtual void putData(bool moption)
	{
		int t_l = convertStrToVal(low);
		int t_leng = convertStrToVal(length);

		int leng = t_leng-t_l+1;

		Out->data.length(leng);
		std::vector<T2> v = mexc->getCellData<T2>(this);
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


#endif