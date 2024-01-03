#pragma once

using namespace System;
using namespace Microsoft::Office::Core;
using namespace Microsoft::Office::Interop::PowerPoint;

namespace ClassLibrary1 {
	public ref class Class1
	{
	public:

		static Microsoft::Office::Interop::PowerPoint::Application^ app;
		static Microsoft::Office::Interop::PowerPoint::Presentations^ presens;
		static Microsoft::Office::Interop::PowerPoint::Presentation^ presen;
		static Microsoft::Office::Interop::PowerPoint::Shape^ tableshape;

		

		static void openPPT(System::String^ path);
		static void savePPT(System::String^ fileName);
		static void closePPT();
		// TODO: このクラスのメソッドをここに追加します。
	};
}
