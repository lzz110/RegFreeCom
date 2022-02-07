// NativeHost.cpp : This file contains the 'main' function. Program execution begins and ends there.
//

#include <iostream>
#include <objbase.h>
#include <comutil.h>

#pragma comment( lib, "comsuppwd.lib")

bool Foo();
bool MyFoo();

int main()
{
    std::cout << "starting...\n";

    HRESULT hr = ::CoInitialize(NULL);

	
	for (int i = 0; i < 3; i++)
	{
		bool succ = Foo();
		if (!succ)
		{
			break;
		}
	}

	for (int i = 0; i < 3; i++)
	{
		bool succ1 = MyFoo();
		if (!succ1)
		{
			break;
		}
	}

    CoFreeUnusedLibraries();
    ::CoUninitialize();

}

bool Foo()
{
	CLSID clsid;
	HRESULT hres;

	HRESULT  hr = CLSIDFromProgID(L"ManagedCom.SomeCom123", &clsid);
	if (hr != S_OK)
	{
		std::cout << "failed 1#" << std::endl;
		return false;
	}
	
	IDispatch* disp1 = NULL;
	hres = CoCreateInstance(clsid, NULL, CLSCTX_INPROC_SERVER, IID_IDispatch, (void**)&disp1);
	if (FAILED(hres))
	{
		std::cout << "failed 2#" << std::endl;
		return false;
	}

	DISPID dispId = -1;
	_bstr_t bt = "GetNextValue";
	BSTR bsMemberName = bt;
	HRESULT	hresult = disp1->GetIDsOfNames(IID_NULL, &bsMemberName, 1, LOCALE_SYSTEM_DEFAULT, &dispId);
	if (hresult != S_OK)
	{
		disp1->Release();
		return false;
	}

	

	DISPPARAMS disparams;
	ZeroMemory(&disparams, sizeof(disparams));
	disparams.cArgs = 0;
	disparams.cNamedArgs = 0;

	VARIANT vt;

	hresult = disp1->Invoke(dispId, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_METHOD, &disparams, &vt, NULL, NULL);
	if (hresult != S_OK)
	{
		std::cout << "failed 3#" << std::endl;
		return false;
	}


	disp1->Release();

	int iRetValue = vt.intVal;

	std::cout << "result:" << iRetValue << std::endl;

	//system("pause");
	return true;
}

bool MyFoo()
{
	CLSID clsid;
	HRESULT hres;

	HRESULT  hr = CLSIDFromProgID(L"TestCom.MyCom123", &clsid);
	if (hr != S_OK)
	{
		std::cout << "failed 1#" << std::endl;
		return false;
	}


	IDispatch* disp1 = NULL;
	hres = CoCreateInstance(clsid, NULL, CLSCTX_INPROC_SERVER, IID_IDispatch, (void**)&disp1);
	if (FAILED(hres))
	{
		std::cout << "failed 2#" << std::endl;
		return false;
	}

	DISPID dispId = -1;
	_bstr_t bt = "GetLastValue";
	BSTR bsMemberName = bt;
	HRESULT	hresult = disp1->GetIDsOfNames(IID_NULL, &bsMemberName, 1, LOCALE_SYSTEM_DEFAULT, &dispId);
	if (hresult != S_OK)
	{
		disp1->Release();
		return false;
	}



	DISPPARAMS disparams;
	ZeroMemory(&disparams, sizeof(disparams));
	disparams.cArgs = 0;
	disparams.cNamedArgs = 0;

	VARIANT vt;

	hresult = disp1->Invoke(dispId, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_METHOD, &disparams, &vt, NULL, NULL);
	if (hresult != S_OK)
	{
		std::cout << "failed 3#" << std::endl;
		return false;
	}


	disp1->Release();

	int iRetValue = vt.intVal;

	std::cout << "result:" << iRetValue << std::endl;

	//system("pause");
	return true;
}

