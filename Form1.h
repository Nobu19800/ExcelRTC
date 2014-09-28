// -*-C++-*-
/*!
 * @file  Form1.h
 * @brief フォーム
 */
#pragma once

#include <vector>
#include <string>
#include <iostream>
#include "myExcel.h"
#include "MyObject.h"


namespace ExcelRTC {

	using namespace System;
	using namespace System::ComponentModel;
	using namespace System::Collections;
	using namespace System::Windows::Forms;
	using namespace System::Data;
	using namespace System::Drawing;

	/// <summary>
	/// Form1 の概要
	///
	/// 警告: このクラスの名前を変更する場合、このクラスが依存するすべての .resx ファイルに関連付けられた
	///          マネージ リソース コンパイラ ツールに対して 'Resource File Name' プロパティを
	///          変更する必要があります。この変更を行わないと、
	///          デザイナと、このフォームに関連付けられたローカライズ済みリソースとが、
	///          正しく相互に利用できなくなります。
	/// </summary>
	/**
	* @class Form1
	*@brief
	*/
	public ref class Form1 : public System::Windows::Forms::Form
	{
	public:
		/**
		*@brief コンストラクタ
		*/
		Form1(void)
		{
			InitializeComponent();
			//
			//TODO: ここにコンストラクタ コードを追加します
			//

			currentDirectory = "c:\\";

			InitFunc();

			
		}
		
		/**
		*@brief ツリーを表示する関数
		*/
		void SetTree();
		
		/**
		*@brief ツリーにオブジェクトを追加する関数
		* @param to ツリーオブジェクト
		* @param tn ツリーノード
		*/
		void AddTree(TreeObject *to, System::Windows::Forms::TreeNode^ tn);
		
		/**
		*@brief 初期設定を行う関数
		*/
		void InitFunc();
		
		/**
		*@brief データポートを作成する関数
		*/
		void CreatePort();
		
		/**
		*@brief データポートを削除する関数
		*/
		void DeletePort();
		
		/**
		*@brief データポートのパラメータをGUIに設定する関数
		*/
		void SetPortParam();
		
		/**
		*@brief 更新した列番号を初期化する関数
		*/
		void ResetPort();
		
		/**
		*@brief 全てのデータポートの更新した列番号を初期化する関数
		*/
		void ResetAllPort();
		
		/**
		*@brief Excelファイルを開く関数
		*/
		void OpenFile();
		
		/**
		*@brief インポートのリストを更新する関数
		*/
		void UpdateInPortList();
		
		/**
		*@brief 関連付けしたインポートのリストを更新する関数
		*/
		void UpdateAttachList();
		
		/**
		*@brief データポートを関連付けする関数
		*/
		void AttachInPort();
		
		/**
		*@brief データポートの関連付けを解除する関数
		*/
		void DetachInPort();
		
		/**
		*@brief 作成したデータポートのリストを更新する関数
		*/
		void UpdatePortList();
		
		/**
		*@brief GUIの値を更新する関数
		* @param m_low 行番号
		* @param m_col 列番号
		* @param m_len 行の範囲
		* @param m_sn シート名
		* @param m_state 列を移動する場合はTrue
		*/
		void SetPortParamEx(std::string m_low, int m_col, std::string m_len, std::string m_sn, bool m_state);
		private: System::String^ currentDirectory; /**<　@brief  */
		public:
		System::Windows::Forms::TreeNode^ m_node;
	private: System::Windows::Forms::ComboBox^  InPortcomboBox;
	private: System::Windows::Forms::ComboBox^  AttachcomboBox;
	public: 

	public: 

	private: System::Windows::Forms::Label^  InPortlabel;
	private: System::Windows::Forms::Label^  Portlabel;
	private: System::Windows::Forms::Button^  Attachbutton;
	private: System::Windows::Forms::Button^  Detachbutton;
	private: System::Windows::Forms::CheckBox^  LowcheckBox;
	private: System::Windows::Forms::ComboBox^  PortListcomboBox;







	public: 	static ExcelRTC::Form1 ^m_form;

	protected:
		/// <summary>
		/// 使用中のリソースをすべてクリーンアップします。
		/// </summary>
		~Form1()
		{
			if (components)
			{
				delete components;
			}
		}
	private: System::Windows::Forms::Button^  Resetbutton;
	private: System::Windows::Forms::Button^  AllResetbutton;
	protected: 

	private: System::Windows::Forms::Button^  Addbutton;

	private: System::Windows::Forms::Button^  Deletebutton;

	private: System::Windows::Forms::Button^  NamingServerbutton;

	private: System::Windows::Forms::Button^  Closebutton;

	private: System::Windows::Forms::Button^  Loadbutton;
	private: System::Windows::Forms::TextBox^  ColtextBox;
	private: System::Windows::Forms::TextBox^  InfotextBox;



	private: System::Windows::Forms::TextBox^  NamingServertextBox;
	private: System::Windows::Forms::TextBox^  LentextBox;


	private: System::Windows::Forms::TextBox^  LawtextBox;
	private: System::Windows::Forms::ComboBox^  SheetcomboBox;


	private: System::Windows::Forms::TreeView^  RTCtreeView;

	private: System::Windows::Forms::OpenFileDialog^  openFileDialog1;
	private: System::Windows::Forms::Label^  Collabel;

	private: System::Windows::Forms::Label^  Sheetlabel;

	private: System::Windows::Forms::Label^  Lawlabel;


	private: System::Windows::Forms::Label^  ManeServerlabel;

	protected: 

	private:
		/// <summary>
		/// 必要なデザイナ変数です。
		/// </summary>
		System::ComponentModel::Container ^components;

#pragma region Windows Form Designer generated code
		/// <summary>
		/// デザイナ サポートに必要なメソッドです。このメソッドの内容を
		/// コード エディタで変更しないでください。
		/// </summary>
		void InitializeComponent(void)
		{
			this->Resetbutton = (gcnew System::Windows::Forms::Button());
			this->AllResetbutton = (gcnew System::Windows::Forms::Button());
			this->Addbutton = (gcnew System::Windows::Forms::Button());
			this->Deletebutton = (gcnew System::Windows::Forms::Button());
			this->NamingServerbutton = (gcnew System::Windows::Forms::Button());
			this->Closebutton = (gcnew System::Windows::Forms::Button());
			this->Loadbutton = (gcnew System::Windows::Forms::Button());
			this->ColtextBox = (gcnew System::Windows::Forms::TextBox());
			this->InfotextBox = (gcnew System::Windows::Forms::TextBox());
			this->NamingServertextBox = (gcnew System::Windows::Forms::TextBox());
			this->LentextBox = (gcnew System::Windows::Forms::TextBox());
			this->LawtextBox = (gcnew System::Windows::Forms::TextBox());
			this->SheetcomboBox = (gcnew System::Windows::Forms::ComboBox());
			this->RTCtreeView = (gcnew System::Windows::Forms::TreeView());
			this->openFileDialog1 = (gcnew System::Windows::Forms::OpenFileDialog());
			this->Collabel = (gcnew System::Windows::Forms::Label());
			this->Sheetlabel = (gcnew System::Windows::Forms::Label());
			this->Lawlabel = (gcnew System::Windows::Forms::Label());
			this->ManeServerlabel = (gcnew System::Windows::Forms::Label());
			this->InPortcomboBox = (gcnew System::Windows::Forms::ComboBox());
			this->AttachcomboBox = (gcnew System::Windows::Forms::ComboBox());
			this->InPortlabel = (gcnew System::Windows::Forms::Label());
			this->Portlabel = (gcnew System::Windows::Forms::Label());
			this->Attachbutton = (gcnew System::Windows::Forms::Button());
			this->Detachbutton = (gcnew System::Windows::Forms::Button());
			this->LowcheckBox = (gcnew System::Windows::Forms::CheckBox());
			this->PortListcomboBox = (gcnew System::Windows::Forms::ComboBox());
			this->SuspendLayout();
			// 
			// Resetbutton
			// 
			this->Resetbutton->Location = System::Drawing::Point(273, 211);
			this->Resetbutton->Name = L"Resetbutton";
			this->Resetbutton->Size = System::Drawing::Size(138, 23);
			this->Resetbutton->TabIndex = 0;
			this->Resetbutton->Text = L"列番号初期化";
			this->Resetbutton->UseVisualStyleBackColor = true;
			this->Resetbutton->Click += gcnew System::EventHandler(this, &Form1::Resetbutton_Click);
			// 
			// AllResetbutton
			// 
			this->AllResetbutton->Location = System::Drawing::Point(273, 240);
			this->AllResetbutton->Name = L"AllResetbutton";
			this->AllResetbutton->Size = System::Drawing::Size(138, 23);
			this->AllResetbutton->TabIndex = 1;
			this->AllResetbutton->Text = L"全ての列番号を初期化";
			this->AllResetbutton->UseVisualStyleBackColor = true;
			this->AllResetbutton->Click += gcnew System::EventHandler(this, &Form1::AllResetbutton_Click);
			// 
			// Addbutton
			// 
			this->Addbutton->Location = System::Drawing::Point(21, 298);
			this->Addbutton->Name = L"Addbutton";
			this->Addbutton->Size = System::Drawing::Size(97, 23);
			this->Addbutton->TabIndex = 2;
			this->Addbutton->Text = L"作成";
			this->Addbutton->UseVisualStyleBackColor = true;
			this->Addbutton->Click += gcnew System::EventHandler(this, &Form1::Addbutton_Click);
			// 
			// Deletebutton
			// 
			this->Deletebutton->Location = System::Drawing::Point(124, 298);
			this->Deletebutton->Name = L"Deletebutton";
			this->Deletebutton->Size = System::Drawing::Size(102, 23);
			this->Deletebutton->TabIndex = 3;
			this->Deletebutton->Text = L"削除";
			this->Deletebutton->UseVisualStyleBackColor = true;
			this->Deletebutton->Click += gcnew System::EventHandler(this, &Form1::Deletebutton_Click);
			// 
			// NamingServerbutton
			// 
			this->NamingServerbutton->Location = System::Drawing::Point(273, 182);
			this->NamingServerbutton->Name = L"NamingServerbutton";
			this->NamingServerbutton->Size = System::Drawing::Size(138, 23);
			this->NamingServerbutton->TabIndex = 4;
			this->NamingServerbutton->Text = L"ツリー表示";
			this->NamingServerbutton->UseVisualStyleBackColor = true;
			this->NamingServerbutton->Click += gcnew System::EventHandler(this, &Form1::NamingServerbutton_Click);
			// 
			// Closebutton
			// 
			this->Closebutton->Location = System::Drawing::Point(273, 298);
			this->Closebutton->Name = L"Closebutton";
			this->Closebutton->Size = System::Drawing::Size(138, 23);
			this->Closebutton->TabIndex = 5;
			this->Closebutton->Text = L"閉じる";
			this->Closebutton->UseVisualStyleBackColor = true;
			this->Closebutton->Click += gcnew System::EventHandler(this, &Form1::Closebutton_Click);
			// 
			// Loadbutton
			// 
			this->Loadbutton->Location = System::Drawing::Point(273, 269);
			this->Loadbutton->Name = L"Loadbutton";
			this->Loadbutton->Size = System::Drawing::Size(138, 23);
			this->Loadbutton->TabIndex = 6;
			this->Loadbutton->Text = L"ファイル読み込み";
			this->Loadbutton->UseVisualStyleBackColor = true;
			this->Loadbutton->Click += gcnew System::EventHandler(this, &Form1::Loadbutton_Click);
			// 
			// ColtextBox
			// 
			this->ColtextBox->Location = System::Drawing::Point(273, 45);
			this->ColtextBox->Name = L"ColtextBox";
			this->ColtextBox->Size = System::Drawing::Size(138, 19);
			this->ColtextBox->TabIndex = 7;
			this->ColtextBox->TextChanged += gcnew System::EventHandler(this, &Form1::ColtextBox_TextChanged);
			// 
			// InfotextBox
			// 
			this->InfotextBox->Location = System::Drawing::Point(273, 7);
			this->InfotextBox->Name = L"InfotextBox";
			this->InfotextBox->Size = System::Drawing::Size(138, 19);
			this->InfotextBox->TabIndex = 8;
			this->InfotextBox->TextChanged += gcnew System::EventHandler(this, &Form1::InfotextBox_TextChanged);
			// 
			// NamingServertextBox
			// 
			this->NamingServertextBox->Location = System::Drawing::Point(273, 156);
			this->NamingServertextBox->Name = L"NamingServertextBox";
			this->NamingServertextBox->Size = System::Drawing::Size(138, 19);
			this->NamingServertextBox->TabIndex = 9;
			this->NamingServertextBox->TextChanged += gcnew System::EventHandler(this, &Form1::NamingServertextBox_TextChanged);
			// 
			// LentextBox
			// 
			this->LentextBox->Location = System::Drawing::Point(352, 81);
			this->LentextBox->Name = L"LentextBox";
			this->LentextBox->Size = System::Drawing::Size(59, 19);
			this->LentextBox->TabIndex = 10;
			// 
			// LawtextBox
			// 
			this->LawtextBox->Location = System::Drawing::Point(273, 81);
			this->LawtextBox->Name = L"LawtextBox";
			this->LawtextBox->Size = System::Drawing::Size(64, 19);
			this->LawtextBox->TabIndex = 11;
			this->LawtextBox->TextChanged += gcnew System::EventHandler(this, &Form1::LawtextBox_TextChanged);
			// 
			// SheetcomboBox
			// 
			this->SheetcomboBox->FormattingEnabled = true;
			this->SheetcomboBox->Location = System::Drawing::Point(273, 118);
			this->SheetcomboBox->Name = L"SheetcomboBox";
			this->SheetcomboBox->Size = System::Drawing::Size(138, 20);
			this->SheetcomboBox->TabIndex = 12;
			this->SheetcomboBox->SelectedIndexChanged += gcnew System::EventHandler(this, &Form1::SheetcomboBox_SelectedIndexChanged);
			// 
			// RTCtreeView
			// 
			this->RTCtreeView->Location = System::Drawing::Point(12, 3);
			this->RTCtreeView->Name = L"RTCtreeView";
			this->RTCtreeView->Size = System::Drawing::Size(255, 260);
			this->RTCtreeView->TabIndex = 13;
			this->RTCtreeView->AfterSelect += gcnew System::Windows::Forms::TreeViewEventHandler(this, &Form1::RTCtreeView_AfterSelect);
			// 
			// openFileDialog1
			// 
			this->openFileDialog1->FileName = L"openFileDialog1";
			// 
			// Collabel
			// 
			this->Collabel->AutoSize = true;
			this->Collabel->Location = System::Drawing::Point(319, 67);
			this->Collabel->Name = L"Collabel";
			this->Collabel->Size = System::Drawing::Size(41, 12);
			this->Collabel->TabIndex = 14;
			this->Collabel->Text = L"行番号";
			// 
			// Sheetlabel
			// 
			this->Sheetlabel->AutoSize = true;
			this->Sheetlabel->Location = System::Drawing::Point(319, 103);
			this->Sheetlabel->Name = L"Sheetlabel";
			this->Sheetlabel->Size = System::Drawing::Size(45, 12);
			this->Sheetlabel->TabIndex = 15;
			this->Sheetlabel->Text = L"シート名";
			// 
			// Lawlabel
			// 
			this->Lawlabel->AutoSize = true;
			this->Lawlabel->Location = System::Drawing::Point(319, 30);
			this->Lawlabel->Name = L"Lawlabel";
			this->Lawlabel->Size = System::Drawing::Size(41, 12);
			this->Lawlabel->TabIndex = 16;
			this->Lawlabel->Text = L"列番号";
			this->Lawlabel->Click += gcnew System::EventHandler(this, &Form1::Lawlabel_Click);
			// 
			// ManeServerlabel
			// 
			this->ManeServerlabel->AutoSize = true;
			this->ManeServerlabel->Location = System::Drawing::Point(281, 141);
			this->ManeServerlabel->Name = L"ManeServerlabel";
			this->ManeServerlabel->Size = System::Drawing::Size(121, 12);
			this->ManeServerlabel->TabIndex = 17;
			this->ManeServerlabel->Text = L"ネームサーバーのアドレス";
			// 
			// InPortcomboBox
			// 
			this->InPortcomboBox->FormattingEnabled = true;
			this->InPortcomboBox->Location = System::Drawing::Point(446, 24);
			this->InPortcomboBox->Name = L"InPortcomboBox";
			this->InPortcomboBox->Size = System::Drawing::Size(121, 20);
			this->InPortcomboBox->TabIndex = 18;
			this->InPortcomboBox->SelectedIndexChanged += gcnew System::EventHandler(this, &Form1::comboBox1_SelectedIndexChanged);
			// 
			// AttachcomboBox
			// 
			this->AttachcomboBox->FormattingEnabled = true;
			this->AttachcomboBox->Location = System::Drawing::Point(446, 67);
			this->AttachcomboBox->Name = L"AttachcomboBox";
			this->AttachcomboBox->Size = System::Drawing::Size(121, 20);
			this->AttachcomboBox->TabIndex = 19;
			this->AttachcomboBox->SelectedIndexChanged += gcnew System::EventHandler(this, &Form1::comboBox2_SelectedIndexChanged);
			// 
			// InPortlabel
			// 
			this->InPortlabel->AutoSize = true;
			this->InPortlabel->Location = System::Drawing::Point(478, 8);
			this->InPortlabel->Name = L"InPortlabel";
			this->InPortlabel->Size = System::Drawing::Size(51, 12);
			this->InPortlabel->TabIndex = 20;
			this->InPortlabel->Text = L"インポート";
			// 
			// Portlabel
			// 
			this->Portlabel->AutoSize = true;
			this->Portlabel->Location = System::Drawing::Point(460, 52);
			this->Portlabel->Name = L"Portlabel";
			this->Portlabel->Size = System::Drawing::Size(99, 12);
			this->Portlabel->TabIndex = 21;
			this->Portlabel->Text = L"関連付けしたInPort";
			// 
			// Attachbutton
			// 
			this->Attachbutton->Location = System::Drawing::Point(459, 103);
			this->Attachbutton->Name = L"Attachbutton";
			this->Attachbutton->Size = System::Drawing::Size(96, 23);
			this->Attachbutton->TabIndex = 22;
			this->Attachbutton->Text = L"関連付け";
			this->Attachbutton->UseVisualStyleBackColor = true;
			this->Attachbutton->Click += gcnew System::EventHandler(this, &Form1::Attachbutton_Click);
			// 
			// Detachbutton
			// 
			this->Detachbutton->Location = System::Drawing::Point(461, 133);
			this->Detachbutton->Name = L"Detachbutton";
			this->Detachbutton->Size = System::Drawing::Size(94, 23);
			this->Detachbutton->TabIndex = 23;
			this->Detachbutton->Text = L"関連付け解除";
			this->Detachbutton->UseVisualStyleBackColor = true;
			this->Detachbutton->Click += gcnew System::EventHandler(this, &Form1::Detachbutton_Click);
			// 
			// LowcheckBox
			// 
			this->LowcheckBox->AutoSize = true;
			this->LowcheckBox->Location = System::Drawing::Point(453, 182);
			this->LowcheckBox->Name = L"LowcheckBox";
			this->LowcheckBox->Size = System::Drawing::Size(106, 16);
			this->LowcheckBox->TabIndex = 24;
			this->LowcheckBox->Text = L"列を移動させるか";
			this->LowcheckBox->UseVisualStyleBackColor = true;
			this->LowcheckBox->CheckedChanged += gcnew System::EventHandler(this, &Form1::LowcheckBox_CheckedChanged);
			// 
			// PortListcomboBox
			// 
			this->PortListcomboBox->FormattingEnabled = true;
			this->PortListcomboBox->Location = System::Drawing::Point(12, 269);
			this->PortListcomboBox->Name = L"PortListcomboBox";
			this->PortListcomboBox->Size = System::Drawing::Size(255, 20);
			this->PortListcomboBox->TabIndex = 25;
			this->PortListcomboBox->SelectedIndexChanged += gcnew System::EventHandler(this, &Form1::PortListcomboBox_SelectedIndexChanged);
			// 
			// Form1
			// 
			this->AutoScaleDimensions = System::Drawing::SizeF(6, 12);
			this->AutoScaleMode = System::Windows::Forms::AutoScaleMode::Font;
			this->ClientSize = System::Drawing::Size(591, 333);
			this->Controls->Add(this->PortListcomboBox);
			this->Controls->Add(this->LowcheckBox);
			this->Controls->Add(this->Detachbutton);
			this->Controls->Add(this->Attachbutton);
			this->Controls->Add(this->Portlabel);
			this->Controls->Add(this->InPortlabel);
			this->Controls->Add(this->AttachcomboBox);
			this->Controls->Add(this->InPortcomboBox);
			this->Controls->Add(this->ManeServerlabel);
			this->Controls->Add(this->Lawlabel);
			this->Controls->Add(this->Sheetlabel);
			this->Controls->Add(this->Collabel);
			this->Controls->Add(this->RTCtreeView);
			this->Controls->Add(this->SheetcomboBox);
			this->Controls->Add(this->LawtextBox);
			this->Controls->Add(this->LentextBox);
			this->Controls->Add(this->NamingServertextBox);
			this->Controls->Add(this->InfotextBox);
			this->Controls->Add(this->ColtextBox);
			this->Controls->Add(this->Loadbutton);
			this->Controls->Add(this->Closebutton);
			this->Controls->Add(this->NamingServerbutton);
			this->Controls->Add(this->Deletebutton);
			this->Controls->Add(this->Addbutton);
			this->Controls->Add(this->AllResetbutton);
			this->Controls->Add(this->Resetbutton);
			this->Name = L"Form1";
			this->Text = L"ExcelRTC";
			this->Load += gcnew System::EventHandler(this, &Form1::Form1_Load);
			this->ResumeLayout(false);
			this->PerformLayout();

		}
#pragma endregion
	private: System::Void NamingServerbutton_Click(System::Object^  sender, System::EventArgs^  e) {
				 SetTree();
			 }
private: System::Void AllResetbutton_Click(System::Object^  sender, System::EventArgs^  e) {
			 ResetAllPort();
		 }
private: System::Void Loadbutton_Click(System::Object^  sender, System::EventArgs^  e) {
			 OpenFile();
		 }
private: System::Void Deletebutton_Click(System::Object^  sender, System::EventArgs^  e) {
			 DeletePort();
		 }
private: System::Void Addbutton_Click(System::Object^  sender, System::EventArgs^  e) {
			 CreatePort();
		 }
private: System::Void Closebutton_Click(System::Object^  sender, System::EventArgs^  e) {
			myExcel::Obj->Close();
		 }
private: System::Void Resetbutton_Click(System::Object^  sender, System::EventArgs^  e) {
			 ResetPort();
		 }
private: System::Void InfotextBox_TextChanged(System::Object^  sender, System::EventArgs^  e) {
		 }
private: System::Void Lawlabel_Click(System::Object^  sender, System::EventArgs^  e) {
		 }
private: System::Void ColtextBox_TextChanged(System::Object^  sender, System::EventArgs^  e) {
		 }
private: System::Void SheetcomboBox_SelectedIndexChanged(System::Object^  sender, System::EventArgs^  e) {
		 }
private: System::Void NamingServertextBox_TextChanged(System::Object^  sender, System::EventArgs^  e) {
		 }
private: System::Void LawtextBox_TextChanged(System::Object^  sender, System::EventArgs^  e) {
		 }
private: System::Void RTCtreeView_AfterSelect(System::Object^  sender, System::Windows::Forms::TreeViewEventArgs^  e) {
			 SetPortParam();
			 PortListcomboBox->Text = "";
		 }
private: System::Void Form1_Load(System::Object^  sender, System::EventArgs^  e) {
		 }
private: System::Void comboBox1_SelectedIndexChanged(System::Object^  sender, System::EventArgs^  e) {
		 }
private: System::Void comboBox2_SelectedIndexChanged(System::Object^  sender, System::EventArgs^  e) {
		 }
private: System::Void Attachbutton_Click(System::Object^  sender, System::EventArgs^  e) {
			 AttachInPort();
		 }
private: System::Void Detachbutton_Click(System::Object^  sender, System::EventArgs^  e) {
			DetachInPort();
		 }
private: System::Void LowcheckBox_CheckedChanged(System::Object^  sender, System::EventArgs^  e) {
			 
		 }
private: System::Void PortListcomboBox_SelectedIndexChanged(System::Object^  sender, System::EventArgs^  e) {
			 SetPortParam();
		 }
};
}

