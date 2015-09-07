// -*- C++ -*-
/*!
 * @file  ExcelTask.h
 * @brief Excelの起動、COMの参照をするタスク
 * $Id$
 */


#ifndef EXCELTASK_H
#define EXCELTASK_H

#include <rtm/RTC.h>



#include <rtm/Manager.h>
#include <rtm/PeriodicExecutionContext.h>

#include "ExcelControl.h"


class ExcelControl;
/*!
 * @class ExcelTask
 * @brief Excelの起動、COMの参照をするスレッド
 *
 */
class ExcelTask : public virtual coil::Task
{
public:
		/**
		*@brief コンストラクタ
		*/
		ExcelTask(ExcelControl *m_rtc);
		/**
		*@brief スレッドを実行
		* @return 
		*/
		virtual int svc();
		
		ExcelControl *mrtc;	/**<　@brief  */


};


#endif