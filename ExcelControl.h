// -*-C++-*-
/*!
 * @file  ExcelControl.h
 * @brief ExcelControlComponent
 */
#ifndef EXCELCONTROL_H
#define EXCELCONTROL_H

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

#include "SubFunction.h"
#ifdef CREATE_EXE_EXCELRTC
	#include "Form1.h"
#else
	#include "ExcelObject.h"
#endif


using namespace RTC;

#include "SpreadSheetSVC_impl.h"
#include "ExcelTask.h"
#include "ExcelDataPort.h"


class ExcelTask;
class ExcelControl;
class ExcelPortBase;
class OtherPort;




/**
*@brief 現在のパスからRTCを検索して、データポートをリストに追加
* @param context ネーミングコンテキスト
* @param rtclist データポートのリスト
* @param name 現在のパス
* @param to RTCのデータポートのツリー
*/
void ListRecursive(CosNaming::NamingContext_ptr context,std::vector<OtherPort> &rtclist,std::string &name, TreeObject *to);

/**
*@brief データポートのリストを作成
* @param naming ネーミングコンテキスト
* @param rtclist データポートのリスト
* @param to RTCのデータポートのツリー
* @param IP_adress IPアドレス
*/
int rtc_get_rtclist(RTC::CorbaNaming &naming,std::vector<OtherPort> &rtclist, TreeObject *to, std::string IP_adress);










/**
*@brief データポートを接続
* @param p1 接続するデータポート
* @param p2 接続するデータポート
*/
void portConnect(PortService_ptr p1, PortService_var p2);



/**
 * @class ExcelControl
*@brief Excelを操作するRTコンポーネント
*/
class ExcelControl
  : public RTC::DataFlowComponentBase
{
 public:
  
	/**
	*@brief コンストラクタ
	* @param manager マネージャオブジェクト
	*/
  ExcelControl(RTC::Manager* manager);

  /**
	*@brief デストラクタ
	*/
  ~ExcelControl();

   
  /**
  *@brief 初期化処理用コールバック関数
  * @return RTC::ReturnCode_t
  */
   virtual RTC::ReturnCode_t onInitialize();

   
   /**
   *@brief 周期処理用コールバック関数
   * @param ec_id target ExecutionContext Id
   * @return RTC::ReturnCode_t
   */
   virtual RTC::ReturnCode_t onExecute(RTC::UniqueId ec_id);
   
   
   /**
   *@brief 不活性化時のコールバック関数
   * @param ec_id target ExecutionContext Id
   * @return RTC::ReturnCode_t
   */
   virtual RTC::ReturnCode_t onDeactivated(RTC::UniqueId ec_id);

   /**
   *@brief 終了処理のコールバック関数
   * @return RTC::ReturnCode_t
   */
   virtual RTC::ReturnCode_t onFinalize();

   
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
   ExcelPortBase* createPort(OtherPort &op, int c, std::string l, std::string sn, std::string leng, bool mstate);
   
   /**
   *@brief データポートを削除する関数
   * @param op 削除するデータポート
   */
   void deleteOtherPort(OtherPort &op);
   /**
   *@brief 全てのデータポートを削除する関数
   */
   void deleteAllPort();
   
   /**
   *@brief RTCの情報を保存用シートに書き込む関数
   */
   void save();
   
   /**
   *@brief RTCの情報をほ場用シートより読み込む関数
   */
   void load();
   
   /**
   *@brief データを書き込む列を初期化する
   * @param mpb データポートオブジェクト
   */
   void resetPort(ExcelPortBase* mpb);
   
   /**
   *@brief 全てのデータを書き込む列を初期化する
   */
   void resetAllPort();
   
   /**
   *@brief コンフィギュレーションパラメータが変更されたときに呼び出される関数
   */
   void configUpdate();

   /**
   *@brief セルの名前を記入
   */
   void update_cellName();


  
   /**
   *@brief 関連付けしたデータポートの処理
   * @param ip インポート
   */
   void updateAPort(ExcelPortBase* ip);

   
   /**
   *@brief セルに値を書き込む関数
   * @param dt 書き込むデータ
   * @param pb データポートオブジェクト
   */
    template <typename T>
   void setCellData(std::vector<std::vector<T>> dt, ExcelPortBase *pb)
	{
		
		ExcelObject::Obj->setCellValue<T>(pb->col+pb->num, pb->low, pb->sheetName, pb->state, dt);
		
	}
   
   /**
   *@brief セルの値を取得する関数
   * @param pb データポートオブジェクト
   * @return 取得したデータ
   */
    template <typename T>
	std::vector<T> getCellData(ExcelPortBase *pb)
	{
		std::vector<T> td = ExcelObject::Obj->getCellValue<T>(pb->col+pb->num, pb->low, pb->sheetName, pb->length);
		
		return td;
	}
	
	/**
	*@brief シーケンス型のデータポートを作成する関数
	* @param op 接続するデータポート
	* @param tdt データ型
    * @param c 列番号
	* @param l 行番号
	* @param sn シート名
	* @param leng 行の範囲
	* @param mstate 列を移動する場合はTrue
	* @return データポートオブジェクト
	*/
	template <typename T, typename T2> ExcelPortBase* crPortSeq(OtherPort &op, std::string tdt, int c, std::string l, std::string sn, std::string leng, bool mstate)
   {
		string PortType = NVUtil::toString(op.pb->get_port_profile()->properties,"port.port_type");
		
	   string tname = op.buff[op.buff.size()-2];
		tname += op.buff[op.buff.size()-1];
	  
		if(PortType == "DataInPort")
		{
			T *m_out = new T();
			OutPort<T> *m_outOut = new OutPort<T>(tname.c_str(),*m_out);
			addOutPort(tname.c_str(), *m_outOut);
			
			ExcelPortBase *mip;
			
			mip= new ExcelOutPortSeq<T,T2>(m_out,m_outOut,tname,tdt,this);
			mip->mop = &op;
			
			mip->setParam(c,l,sn,leng,mstate);
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

			ExcelPortBase *mip = new ExcelInPortSeq<T,T2>(m_in,m_inIn,tname,tdt,this);
			mip->mop = &op;
			

			mip->setParam(c,l,sn,leng,mstate);
			InPorts.push_back(mip);
			
			portConnect(m_inIn->getPortRef(), op.pb);

			op.mpb = mip;

			return mip;
		}
		return NULL;
	}
	
	/**
	*@brief データポートを作成する関数
	* @param op 接続するデータポート
	* @param tdt データ型
    * @param c 列番号
	* @param l 行番号
	* @param sn シート名
	* @param leng 行の範囲
	* @param mstate 列を移動する場合はTrue
	* @return データポートオブジェクト
	*/
   template <typename T, typename T2> ExcelPortBase* crPort(OtherPort &op, std::string tdt, int c, std::string l, std::string sn, std::string leng, bool mstate)
   {
	   string PortType = NVUtil::toString(op.pb->get_port_profile()->properties,"port.port_type");
		
	   string tname = op.buff[op.buff.size()-2];
		tname += op.buff[op.buff.size()-1];
	  
		if(PortType == "DataInPort")
		{
			T *m_out = new T();
			OutPort<T> *m_outOut = new OutPort<T>(tname.c_str(),*m_out);
			addOutPort(tname.c_str(), *m_outOut);
			
			ExcelPortBase *mip;
			
			mip= new ExcelOutPort<T,T2>(m_out,m_outOut,tname,tdt,this);
			mip->mop = &op;
			
			mip->setParam(c,l,sn,leng,mstate);
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

			ExcelPortBase *mip = new ExcelInPort<T,T2>(m_in,m_inIn,tname,tdt,this);
			mip->mop = &op;
			

			mip->setParam(c,l,sn,leng,mstate);
			InPorts.push_back(mip);
			
			portConnect(m_inIn->getPortRef(), op.pb);

			op.mpb = mip;

			return mip;
		}
		return NULL;
   }
   
   /**
   *@brief コンフィギュレーションパラメータによりシーケンス型のデータポートを作成する関数
    * @param tname 名称
	* @param PortType データポートのタイプ
	* @param tdt データ型
	* @param c 列番号
	* @param l 行番号
	* @param sn シート名
	* @param leng 行の範囲
	* @param mstate 列を移動する場合はTrue
	* @return データポートオブジェクト
   */
   template <typename T, typename T2> ExcelPortBase* confcrPortSeq(std::string tname, std::string PortType, std::string tdt, int c, std::string l, std::string sn, std::string leng, bool mstate)
   {
		
	  
		if(PortType == "DataOutPort")
		{
			T *m_out = new T();
			OutPort<T> *m_outOut = new OutPort<T>(tname.c_str(),*m_out);
			addOutPort(tname.c_str(), *m_outOut);
			
			ExcelPortBase *mip;
			
			mip= new ExcelOutPortSeq<T,T2>(m_out,m_outOut,tname,tdt,this);
			
			mip->setParam(c,l,sn,leng,mstate);
			ConfOutPorts.push_back(mip);
			

			

			return mip;
		
		}
		else if(PortType == "DataInPort")
		{
			T *m_in = new T();
			InPort<T> *m_inIn = new InPort<T>(tname.c_str(),*m_in);

			addInPort(tname.c_str(), *m_inIn);

			ExcelPortBase *mip = new ExcelInPortSeq<T,T2>(m_in,m_inIn,tname,tdt,this);
			
			

			mip->setParam(c,l,sn,leng,mstate);
			ConfInPorts.push_back(mip);
			

			

			return mip;
		}
		return NULL;
	}
	
   /**
   *@brief コンフィギュレーションパラメータによりデータポートを作成する関数
    * @param tname 名称
	* @param PortType データポートのタイプ
	* @param tdt データ型
	* @param c 列番号
	* @param l 行番号
	* @param sn シート名
	* @param leng 行の範囲
	* @param mstate 列を移動する場合はTrue
	* @return データポートオブジェクト
   */
   template <typename T, typename T2> ExcelPortBase* confcrPort(std::string tname, std::string PortType, std::string tdt, int c, std::string l, std::string sn, std::string leng, bool mstate)
   {
	  
	  
		if(PortType == "DataOutPort")
		{
			T *m_out = new T();
			OutPort<T> *m_outOut = new OutPort<T>(tname.c_str(),*m_out);
			addOutPort(tname.c_str(), *m_outOut);
			
			ExcelPortBase *mip;
			
			mip= new ExcelOutPort<T,T2>(m_out,m_outOut,tname,tdt,this);
			
			mip->setParam(c,l,sn,leng,mstate);
			ConfOutPorts.push_back(mip);
			

			

			return mip;
		
		}
		else if(PortType == "DataInPort")
		{
			T *m_in = new T();
			InPort<T> *m_inIn = new InPort<T>(tname.c_str(),*m_in);
			

			addInPort(tname.c_str(), *m_inIn);

			ExcelPortBase *mip = new ExcelInPort<T,T2>(m_in,m_inIn,tname,tdt,this);
			
			

			mip->setParam(c,l,sn,leng,mstate);
			ConfInPorts.push_back(mip);
			

			

			return mip;
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
   ExcelPortBase* setDPort(std::vector<std::string> pt, int c, std::string l, std::string sn, std::string leng, bool mstate, bool msflag);
   
   /**
   *@brief データポートを削除する関数
   * @param pt 削除するデータポートのパス
   */
   void delDPort(std::vector<std::string> pt);
   
   /**
   *@brief データポートを取得する関数
   * @param pt 取得するデータポートのパス
   * @return データポートオブジェクト
   */
   ExcelPortBase* getDPort(std::vector<std::string> pt);
   
   
   /**
   *@brief データポートを関連付ける関数
   * @param mpb 関連付けるアウトポート
   * @param n 関連付けるインポートの名前
   */
   void attachPort(ExcelPortBase *mpb, std::string n);
   
   /**
   *@brief データポートの関連付けを解除する関数
   * @param mpb 関連付けを解除するアウトポート
   * @param n 関連付けを解除するインポートの名前
   */
   void detachPort(ExcelPortBase *mpb, std::string n);
   
   /**
   *@brief インポートを取得する関数
   * @param n インポートの名前
   * @return データポートオブジェクト
   */
   ExcelPortBase *getInPort(std::string n);
   
   /**
   *@brief アウトポートを取得する関数
   * @param n アウトポートの名前
   * @return データポートオブジェクト
   */
   ExcelPortBase *getOutPort(std::string n);
   
   /**
   *@brief コンフィギュレーションパラメータで設定したアウトポートを取得する関数
   * @param n アウトポートの名前
   * @return データポートオブジェクト
   */
   ExcelPortBase *getConfOutPort(std::string n);
   /**
   *@brief コンフィギュレーションパラメータで設定したインポートを取得する関数
   * @param n インポートの名前
   * @return データポートオブジェクト
   */
   ExcelPortBase *getConfInPort(std::string n);

   
   /**
   *@brief RTCのデータポートのツリーを取得する関数
   * @param IP_adress ネームサーバの名前
   * @return ツリーオブジェクト
   */
   TreeObject* getRTCTree(std::string IP_adress);

   /**
   *@brief ファイル名のコンフィギュレーションパラメータ変更の関数
   * @param FP ファイル名
   */
   void setFilePath(std::string FP);

   std::vector<ExcelPortBase*> InPorts;	/**<　@brief  */
   std::vector<ExcelPortBase*> OutPorts;	/**<　@brief  */

   std::vector<ExcelPortBase*> ConfInPorts;	/**<　@brief  */
   std::vector<ExcelPortBase*> ConfOutPorts;	/**<　@brief  */

   RTC::Manager* m_manager;	/**<　@brief  */

   std::vector<OtherPort> rtclist;	/**<　@brief  */

   RTC::CorbaPort m_SpreadSheetPort;	/**<　@brief  */
   mSpreadSheetSVC_impl m_spreadsheet;	/**<　@brief  */


   /**
   *@brief ファイルを取得する関数
   * @return ファイル名
   */
   std::string getFileName();

 protected:
	std::string file_path;	/**<　@brief  */
	std::string conf_data_type;	/**<　@brief  */
	std::string conf_port_type;	/**<　@brief  */
	int conf_column;	/**<　@brief  */
	std::string conf_start_row;	/**<　@brief  */
	std::string conf_end_row;	/**<　@brief  */
	std::string conf_sheetname;	/**<　@brief  */
	int c_move;
	std::string attach_Port;

	
	

	int actionLock;	/**<　@brief  */
	int red;	/**<　@brief  */
	int green;	/**<　@brief  */
	int blue;	/**<　@brief  */

	coil::Mutex _mutex; /**<　@brief  */

	ExcelTask *et;


 private:


};

/**
 * @class ExcelConfigUpdateParam
*@brief コンフィギュレーションパラメータが更新されたときのコールバック
*/
class ExcelConfigUpdateParam
    : public RTC::ConfigurationSetListener
{
public:
	/**
	*@brief コンストラクタ
	* @param e_rtc ExcelRTC
	*/
    ExcelConfigUpdateParam(ExcelControl *e_rtc)
    {
		m_rtc = e_rtc;
    }
	/**
	*@brief 
	* @param config_set 
	*/
    void operator()(const coil::Properties& config_set)
	{
		
		m_rtc->configUpdate();
		
    }
	ExcelControl *m_rtc; /**<　@brief  */

};

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


extern "C"
{
  DLL_EXPORT void ExcelControlInit(RTC::Manager* manager);
};

#endif