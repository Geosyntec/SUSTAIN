//-----------------------------------------------------------------------------
//   swmm5.h
//
//   Project: EPA SWMM5
//   Version: 5.0
//   Date:    5/6/05   (Build 5.0.005)
//   Author:  L. Rossman
//
//   Prototypes for SWMM5 functions exported to swmm5.dll.
//-----------------------------------------------------------------------------

#ifdef DLL
	#ifdef __cplusplus
		#define DLLEXPORT extern "C" __declspec(dllexport) __stdcall
	#else
		#define DLLEXPORT __declspec(dllexport) __stdcall
	#endif
#else
	#define DLLEXPORT int
#endif

DLLEXPORT swmm_run(char* f1, char* f2, char* f3);
DLLEXPORT swmm_open(char* f1, char* f2, char* f3);
DLLEXPORT swmm_start(int saveFlag);
DLLEXPORT swmm_step(double* elapsedTime);
DLLEXPORT swmm_end(void);
DLLEXPORT swmm_report(void);
DLLEXPORT swmm_getMassBalErr(float* runoffErr, float* flowErr,float* qualErr);
DLLEXPORT swmm_close(void);
DLLEXPORT swmm_getVersion(void);

	
