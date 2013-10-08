#include "stdafx.h"

#define SAFE_RELEASE(p) { if((p)!=NULL) { (p)->Release(); (p)=NULL; }}
#define FAILED_RELEASE(r) if(FAILED(r)){ printf("%s:(%d) Failed. HRESULT=%x\n",__FILE__,__LINE__,hr); ReleaseInterfaces(); return 0; }
#define CREATE_COM_INST(clsid,iid,obj) CoCreateInstance(clsid, NULL, CLSCTX_INPROC_SERVER, iid,reinterpret_cast<void **>(obj))

ICaptureGraphBuilder2 *pCapGrph = NULL;
IGraphBuilder *        pGrph = NULL; 
IMediaControl *        pCtrl = NULL;
IBaseFilter *          pDVCam= NULL;		// DVcam を表すフィルタ
IBaseFilter *		   pSplitter  =NULL;
IBaseFilter *		   pRenderer  =NULL;
IBaseFilter *		   pDSRenderer=NULL;
IBaseFilter *		   pDecoder   =NULL;

void ReleaseInterfaces(){
	SAFE_RELEASE(pCapGrph);
	SAFE_RELEASE(pGrph);
	SAFE_RELEASE(pCtrl);
	SAFE_RELEASE(pDVCam);
	SAFE_RELEASE(pSplitter);
	SAFE_RELEASE(pRenderer);
	SAFE_RELEASE(pDSRenderer);
	CoUninitialize();
}

HRESULT CreateDvCam(){
	ICreateDevEnum *pDevEnum = NULL;
	IEnumMoniker *pEnum = NULL;
	// システム デバイス列挙子を作成する。
	HRESULT hr=CREATE_COM_INST(CLSID_SystemDeviceEnum,IID_ICreateDevEnum,&pDevEnum);
	if(FAILED(hr)) {
		return hr;
	}
	// ビデオ キャプチャ カテゴリの列挙子を作成する。
	hr = pDevEnum->CreateClassEnumerator(CLSID_VideoInputDeviceCategory,&pEnum, 0);
	if(hr!=S_OK){
		return hr;
	}
	IMoniker *pMoniker = NULL;
	while (pEnum->Next(1, &pMoniker, NULL) == S_OK && pDVCam==NULL)
	{
		IPropertyBag *pPropBag=NULL;
		hr = pMoniker->BindToStorage(0, 0, IID_IPropertyBag, reinterpret_cast<void**>(&pPropBag));
		if (SUCCEEDED(hr)){
			VARIANT varName;
            VariantInit(&varName);
            hr = pPropBag->Read(TEXT("FriendlyName"), &varName, 0);
            if (SUCCEEDED(hr)){
				wprintf(TEXT("%s\n"),varName.bstrVal);
            }
			// Description を取得できるのは、DV と D-VHS/MPEGのカムコーダデバイスのみ。
			VariantClear(&varName);
            VariantInit(&varName);
            hr = pPropBag->Read(TEXT("Description"), &varName, 0);
            if (SUCCEEDED(hr)){
				wprintf(TEXT("%s\n"),varName.bstrVal);
				// IBaseFilter を取得する
				hr = pMoniker->BindToObject(0, 0, IID_IBaseFilter, reinterpret_cast<void**>(&pDVCam));
			}
			VariantClear(&varName);
		} 
		SAFE_RELEASE(pPropBag);
		SAFE_RELEASE(pMoniker);
	}
	return pDVCam==NULL ? E_FAIL : S_OK;
}

int _tmain(int argc, _TCHAR* argv[]){
	HRESULT hr=NOERROR;
	IBaseFilter *pBaseF=NULL;
	hr=CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
	//  graph
	FAILED_RELEASE(CREATE_COM_INST(CLSID_CaptureGraphBuilder2, IID_ICaptureGraphBuilder2,&pCapGrph));
	FAILED_RELEASE(CREATE_COM_INST(CLSID_FilterGraph, IID_IGraphBuilder,&pGrph));
	FAILED_RELEASE(pCapGrph->SetFiltergraph(pGrph));
	FAILED_RELEASE(pGrph->QueryInterface(IID_IMediaControl, (void**)&pCtrl));
	// add filter : DV cam
	FAILED_RELEASE(CreateDvCam());
	FAILED_RELEASE(pGrph->AddFilter(pDVCam,TEXT("DVCam")));
	// add filter : DV splitter
	FAILED_RELEASE(CREATE_COM_INST(CLSID_DVSplitter, IID_IBaseFilter,&pSplitter));
	FAILED_RELEASE(pGrph->AddFilter(pSplitter,TEXT("Splitter")));
	// add filter : DV decoder
	FAILED_RELEASE(CREATE_COM_INST(CLSID_DVVideoCodec, IID_IBaseFilter, &pDecoder));
	FAILED_RELEASE(pGrph->AddFilter(pDecoder,TEXT("DV Decoder")));
	// set decoder resolution
	IIPDVDec *pDD=NULL;
	FAILED_RELEASE(pDecoder->QueryInterface(IID_IIPDVDec,(void**)&pDD));
	pDD->put_IPDisplay(DVDECODERRESOLUTION_360x240);
	SAFE_RELEASE(pDD);
	// add filter : video renderer
	FAILED_RELEASE(CREATE_COM_INST(CLSID_VideoRenderer, IID_IBaseFilter, &pRenderer));
	FAILED_RELEASE(pGrph->AddFilter(pRenderer,TEXT("Video Renderer")));
	// add filter : directsound renderer
	FAILED_RELEASE(CREATE_COM_INST(CLSID_DSoundRender, IID_IBaseFilter, &pDSRenderer));
	FAILED_RELEASE(pGrph->AddFilter(pDSRenderer,TEXT("DirectSound Renderer")));
	// Connect using ICaptureGraphBuilder2::RenderStream
	FAILED_RELEASE(pCapGrph->RenderStream(&PIN_CATEGORY_CAPTURE ,&MEDIATYPE_Interleaved,pDVCam,NULL,pSplitter));
	FAILED_RELEASE(pCapGrph->RenderStream(NULL,&MEDIATYPE_Video,pSplitter,pDecoder,pRenderer));
	FAILED_RELEASE(pCapGrph->RenderStream(NULL,&MEDIATYPE_Audio,pSplitter,NULL,pDSRenderer));
		
	pCtrl->Run();
	_getch();
	pCtrl->StopWhenReady();
	ReleaseInterfaces();
	return 0;
}
