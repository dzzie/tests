
#include <windows.h>
#include "msvbvm60.tlh"
#include <conio.h>

//to generate the tlh
//#import "C:\\windows\system32\msvbvm60.dll" no_namespace

void addStr(_CollectionPtr p , char* str){
	_variant_t vv;
	vv.SetString(str);
	p->Add( &vv.GetVARIANT() );

	/*VARIANT v;
	VariantInit(&v);
	v.bstrVal = SysAllocString(L"this is string 1!");
	v.vt = VT_BSTR;*/
}

/*
.text:00401638 dword_401638    dd 0                    ; flag can be 0,1,2
.text:0040163C                 dd offset dword_401618  ; {A4C4671C-499F-101B-BB78-00AA00383CBB}  ;Class Collection
.text:00401640                 dd offset dword_401628  ; {A4C46780-499F-101B-BB78-00AA00383CBB}  ;Interface _Collection
                               dd 0                    ; required
*/
struct vbComDef{
	int flag;
	GUID* clsid;
	GUID* IFaceID;
	int unk;
};

//void (*fn_vbaNew)(vbComDef*);
//void (*fn_vbaNew2)(vbComDef*, void*);

typedef void*(__stdcall *fn_vbaNew)(vbComDef*);
typedef void (__stdcall *fn_vbaNew2)(vbComDef*, void*);
typedef void (__stdcall *CreateIExprSrvObj)(int,int,int);

void main(void)
{
	// {A4C4671C-499F-101B-BB78-00AA00383CBB}
	// IID clsid = {0xA4C46780,0x499F,0x101B,0xBB,0x78,0x00,0xAA,0x00,0x38,0x3C,0xBB};

    HMODULE h = LoadLibrary("msvbvm60.dll");

	fn_vbaNew vbaNew = (fn_vbaNew)GetProcAddress(h, "__vbaNew");
	fn_vbaNew2 vbaNew2 = (fn_vbaNew2)GetProcAddress(h, "__vbaNew2");
	CreateIExprSrvObj IExprSrvObj = (CreateIExprSrvObj)GetProcAddress(h,"CreateIExprSrvObj"); 

	if( vbaNew == NULL || vbaNew2 == NULL){
		printf("Failed to load vbaNew/2");
		return;
	}

	GUID clsid;
	GUID ifaceid;
	
	if( FAILED(CLSIDFromString(L"{A4C4671C-499F-101B-BB78-00AA00383CBB}", &clsid))){
		printf("clsid failed");
		return;
	}

	if( FAILED(CLSIDFromString(L"{A4C46780-499F-101B-BB78-00AA00383CBB}", &ifaceid))){
		printf("IFaceID failed");
		return;
	}

	vbComDef cd;
	cd.flag = 0;
	cd.clsid = &clsid;
	cd.IFaceID = &ifaceid;
	cd.unk = 0; //must be set to 0

	CoInitialize(NULL);
	
	IExprSrvObj(0,4,0); // we need TLS initilized in the runtime...
	//this would not be needed if had loaded a vb COM object already but for external testing..

	 _CollectionPtr c; 
	 
	 //_asm int 3
	 vbaNew2(&cd, &c);

  /*
	 void* p=0;
	 p = vbaNew(&cd); //not sure how to cast to a smart pointer normally.. use __vbaObjSet maybe to do the dirty work?
	 _asm{            //or asm works..
		 mov eax, p
		 mov c, eax
	 } 
  */

	VARIANT v;
	VariantInit(&v);
	v.bstrVal = SysAllocString(L"this is string 1!");
	v.vt = VT_BSTR;
	c->Add(&v);

	VARIANT v2;
	VariantInit(&v2);
	v2.bstrVal = SysAllocString(L"this is string 2!");
	v2.vt = VT_BSTR;
	c->Add(&v2);

	addStr(c, "this is my wrapper!");
	addStr(c, "this is my wrapper1");
	addStr(c, "this is my wrapper2");
	addStr(c, "this is my wrapper3");
	
	VARIANT vi;
	VariantInit(&vi);
	vi.vt = VT_I4;

	for(int i=1; i< c->Count(); i++){

		vi.intVal = i;
		VARIANT v = c->Item(&vi);

		if(v.vt == VT_BSTR){
			printf("%d  %ws\n",i, v.bstrVal);
		}
		else{
			printf("%d vt:%x \n", i, v.vt);
		}

	}

     getch();

}