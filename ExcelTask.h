// -*- C++ -*-
/*!
 * @file  ExcelTask.h
 *
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
 * @brief 
 *
 */
class ExcelTask : public virtual coil::Task
{
public:
		/**
		*@brief コンストラクタ
		* @param manager
		*/
		ExcelTask(ExcelControl *m_rtc);
		/**
		*@brief 
		*/
		virtual int svc();
		
		ExcelControl *mrtc;	/**<　@brief  */


};


#endif