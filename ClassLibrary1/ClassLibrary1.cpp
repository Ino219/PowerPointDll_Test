#include "pch.h"

#include "ClassLibrary1.h"

void ClassLibrary1::Class1::openPPT(System::String ^ path)
{
	app = gcnew Microsoft::Office::Interop::PowerPoint::ApplicationClass();
	presens = app->Presentations;
	presen = presens->Open(
		path,
		MsoTriState::msoFalse,
		MsoTriState::msoFalse,
		MsoTriState::msoFalse
	);
}

void ClassLibrary1::Class1::savePPT(System::String ^ fileName)
{
	//�w�肵���t�@�C�����ŕۑ�
	presen->SaveAs(fileName, Microsoft::Office::Interop::PowerPoint::PpSaveAsFileType::ppSaveAsDefault, MsoTriState::msoTrue);

}

void ClassLibrary1::Class1::closePPT()
{
	//���\�[�X�̊J��
	//System::Runtime::InteropServices::Marshal::ReleaseComObject(tableshape);

	presen->Close();
	System::Runtime::InteropServices::Marshal::ReleaseComObject(presen);
	System::Runtime::InteropServices::Marshal::ReleaseComObject(presens);

	app->Quit();
	System::Runtime::InteropServices::Marshal::ReleaseComObject(app);
}
