/*!
 * @file  FormControl.cpp
 * @brief フォームの入力処理
 *
 */

#include "stdafx.h"
#include "Form1.h"
#include <sstream>

#include "ExcelControl.h"

extern ExcelControl *tertc;



void ExcelRTC::Form1::SetTree()
{
	this->RTCtreeView->Nodes->Clear();
	TreeObject *to = tertc->getRTCTree(MarshalString(NamingServertextBox->Text));
	
	if(to)
	{
		m_node = this->RTCtreeView->Nodes->Add("/");
	
		AddTree(to, this->RTCtreeView->Nodes[0]);


		this->SheetcomboBox->Items->Clear();
		for(int i=0;i < ExcelObject::Obj->xlWorksheet->Count;i++)
		{
			this->SheetcomboBox->Items->Add(ExcelObject::Obj->xlWorksheet[i]->Name);
		}
		this->SheetcomboBox->Text = ExcelObject::Obj->xlWorksheet[0]->Name;
	
		LawtextBox->Text = "A";
		ColtextBox->Text = "1";
		LentextBox->Text = "A";
		LowcheckBox->Checked = true;

		tertc->load();
	
		UpdatePortList();
	}
	
	
}


void ExcelRTC::Form1::AttachInPort()
{
	ExcelPortBase *mpb = tertc->getInPort(MarshalString(PortListcomboBox->Text));
	
	if(mpb == NULL)
	{
		mpb = tertc->getOutPort(MarshalString(PortListcomboBox->Text));
		if(mpb == NULL)
		{
			std::vector<std::string> bt = coil::split(MarshalString(this->RTCtreeView->SelectedNode->FullPath), "\\");
			if(bt.size() > 1)
			{
				bt.erase(bt.begin());
				mpb = tertc->getDPort(bt);
			}
		}
	}
	
		
	
		
	if(mpb != NULL)
	{
		tertc->attachPort(mpb, MarshalString(InPortcomboBox->Text));
	}
	else
	{
		System::Windows::Forms::MessageBox::Show("削除済みです");
	}

	UpdateAttachList();
	
}

void ExcelRTC::Form1::DetachInPort()
{
	ExcelPortBase *mpb = tertc->getInPort(MarshalString(PortListcomboBox->Text));
	if(mpb == NULL)
	{
		mpb = tertc->getOutPort(MarshalString(PortListcomboBox->Text));
		if(mpb == NULL)
		{
			std::vector<std::string> bt = coil::split(MarshalString(this->RTCtreeView->SelectedNode->FullPath), "\\");
			if(bt.size() > 1)
			{
				bt.erase(bt.begin());
				mpb = tertc->getDPort(bt);
			}
		}
	}

	if(mpb != NULL)
	{
		tertc->detachPort(mpb, MarshalString(AttachcomboBox->Text));
		AttachcomboBox->Text = "";
	}
	else
	{
		System::Windows::Forms::MessageBox::Show("削除済みです");
	}

	UpdateAttachList();
	
	
}

void ExcelRTC::Form1::UpdateInPortList()
{
	
	InPortcomboBox->Items->Clear();
	for(int i=0;i < tertc->InPorts.size();i++)
	{
		InPortcomboBox->Items->Add(gcnew System::String(tertc->InPorts[i]->name.c_str()));
	}
}


void ExcelRTC::Form1::UpdateAttachList()
{
	ExcelPortBase *mpb = tertc->getInPort(MarshalString(PortListcomboBox->Text));

	if(mpb == NULL)
	{
		mpb = tertc->getOutPort(MarshalString(PortListcomboBox->Text));
		if(mpb == NULL)
		{
			std::vector<std::string> bt = coil::split(MarshalString(this->RTCtreeView->SelectedNode->FullPath), "\\");
			if(bt.size() > 1)
			{
				bt.erase(bt.begin());
				mpb = tertc->getDPort(bt);
			}
		}
	}
	
	
	if(mpb != NULL)
	{
		AttachcomboBox->Items->Clear();
		for(int i=0;i < mpb->attachPort.size();i++)
		{
			AttachcomboBox->Items->Add(gcnew System::String(mpb->attachPort[i].c_str()));
		}
	}
	
	
}

void ExcelRTC::Form1::SetPortParamEx(std::string m_low, int m_col, std::string m_len, std::string m_sn, bool m_state)
{
	
	this->LawtextBox->Text = gcnew System::String(m_low.c_str());
	std::stringstream ss2;
	ss2 << m_col;
	this->ColtextBox->Text = gcnew System::String(ss2.str().c_str());
	
	this->LentextBox->Text = gcnew System::String(m_len.c_str());
	this->SheetcomboBox->Text = gcnew System::String(m_sn.c_str());

	this->LowcheckBox->Checked = m_state;
}

void ExcelRTC::Form1::SetPortParam()
{
	ExcelPortBase *mpb = tertc->getInPort(MarshalString(PortListcomboBox->Text));
	
	if(mpb == NULL)
	{
		mpb = tertc->getOutPort(MarshalString(PortListcomboBox->Text));
		if(mpb == NULL)
		{
			std::vector<std::string> bt = coil::split(MarshalString(this->RTCtreeView->SelectedNode->FullPath), "\\");
			if(bt.size() > 1)
			{
				bt.erase(bt.begin());
				mpb = tertc->getDPort(bt);
			}
		}
	}

	
	
	if(mpb != NULL)
	{
		
			
		SetPortParamEx(mpb->low, mpb->col, mpb->length, mpb->sheetName, mpb->state);

		this->InfotextBox->Text = "作成済み";

		

	}
	else
	{
		this->InfotextBox->Text = "未作成";
	}

	UpdateInPortList();
	UpdateAttachList();
	

	

		
}
void ExcelRTC::Form1::resetPort()
{
	std::vector<std::string> bt = coil::split(MarshalString(this->RTCtreeView->SelectedNode->FullPath), "\\");
	
	if(bt.size() > 1)
	{
		bt.erase(bt.begin());
	

		ExcelPortBase *mpb = tertc->getDPort(bt);
		if(mpb != NULL)
		{
			mpb->num = 0;
		}
	}

}

void ExcelRTC::Form1::resetAllPort()
{
	tertc->resetAllPort();
}

void ExcelRTC::Form1::createPort()
{
		
	int c;
	coil::stringTo<int>(c,MarshalString(this->ColtextBox->Text).c_str());

	std::string l = MarshalString(this->LawtextBox->Text);
	std::string leng = MarshalString(this->LentextBox->Text);
	std::string sn = MarshalString(this->SheetcomboBox->Text);
	bool mstate = LowcheckBox->Checked;

	ExcelPortBase *mpb = tertc->getInPort(MarshalString(PortListcomboBox->Text));
	if(mpb != NULL)
	{
		mpb->setParam(c, l, sn, leng, mstate);
	}
	else
	{
		mpb = tertc->getOutPort(MarshalString(PortListcomboBox->Text));
		if(mpb != NULL)
		{
			mpb->setParam(c, l, sn, leng, mstate);
		}
		else
		{
			std::vector<std::string> bt = coil::split(MarshalString(this->RTCtreeView->SelectedNode->FullPath), "\\");

			if(bt.size() > 1)
			{
				bt.erase(bt.begin());

				
				ExcelPortBase* mpb = tertc->setDPort(bt,c,l,sn,leng,mstate,true);
				if(mpb != NULL)
				{
					this->InfotextBox->Text = "作成済み";
					tertc->save();

					
				}
				else
				{
					this->InfotextBox->Text = "未作成";
					System::Windows::Forms::MessageBox::Show("データポートではありません");
				}
			}
		}
	}

	

	UpdateInPortList();
	UpdateAttachList();
	UpdatePortList();
	
}

void ExcelRTC::Form1::DeletePort()
{
	ExcelPortBase *mpb = tertc->getInPort(MarshalString(PortListcomboBox->Text));
	if(mpb != NULL)
	{
		tertc->deleteOtherPort(*(mpb->mop));
	}
	else
	{
		mpb = tertc->getOutPort(MarshalString(PortListcomboBox->Text));
		if(mpb != NULL)
		{
			tertc->deleteOtherPort(*(mpb->mop));
		}
		else
		{
			std::vector<std::string> bt = coil::split(MarshalString(this->RTCtreeView->SelectedNode->FullPath), "\\");
			bt.erase(bt.begin());
	
			tertc->delDPort(bt);
		}
	}

	this->InfotextBox->Text = "未作成";

	UpdateInPortList();
	UpdateAttachList();
	UpdatePortList();

	PortListcomboBox->Text = "";
}

void ExcelRTC::Form1::AddTree(TreeObject *to, System::Windows::Forms::TreeNode^ tn)
{
	System::Windows::Forms::TreeNode^ ntn = tn->Nodes->Add(gcnew System::String(to->name.c_str()));
	
	for(int i=0;i < to->to.size();i++)
	{
		AddTree(to->to[i], ntn);
	}
}


void ExcelRTC::Form1::InitFunc()
{
	m_form = this;
	NamingServertextBox->Text = "localhost:2809";

	if(tertc)
	{
		tertc->load();
		
		UpdatePortList();
	}
}

void ExcelRTC::Form1::OpenFile()
{
	
	openFileDialog1->Title = "ファイル開くダイアログボックス";
	 openFileDialog1->InitialDirectory = currentDirectory;
	 openFileDialog1->Filter = "EXCELファイル(*.xls;*.xlsx)|*.xls;*.xlsx|すべてのファイル(*.*)|*.*";

	 if(openFileDialog1->ShowDialog() == System::Windows::Forms::DialogResult::OK)
	 {
		 std::string filePath = MarshalString(openFileDialog1->FileName);
		 coil::replaceString(filePath, "\\", "/");
		 tertc->setFilePath(filePath);
		 ExcelObject::Obj->Open(openFileDialog1->FileName);
		 currentDirectory = System::IO::Path::GetDirectoryName(openFileDialog1->FileName);
		 //tertc->load();
		 
		 UpdatePortList();
	 }
}

void ExcelRTC::Form1::UpdatePortList()
{
	//MarshalString(AttachcomboBox->Text)

	
	PortListcomboBox->Items->Clear();
	for(int i=0;i < tertc->InPorts.size();i++)
	{
		PortListcomboBox->Items->Add(gcnew System::String(tertc->InPorts[i]->name.c_str()));
	}
	for(int i=0;i < tertc->OutPorts.size();i++)
	{
		PortListcomboBox->Items->Add(gcnew System::String(tertc->OutPorts[i]->name.c_str()));
	}
}