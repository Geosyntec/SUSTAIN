//-----------------------------------------------------------------------------
//   Sediment.h
//
//   Project:  EPA SUSTAIN
//   Version:  1.0
//   Date:     3/20/07   
//

float dayval(float mval1,float mval2,int day,int ndays);
int detach(int crvfg,int csnofg,int mon,int nxtmon,int day,int ndays,
		   float *coverm,float rain,float snocov,float delt60,float smpf, 
		   float krer,float jrer,float& cover,float& dets,float& det);
int attach(float affix,float deltd,float& dets);
int sosed1(float runoff,float surs,float delt60,float kser,float jser,
		   float kger,float jger,float& dets,float& sosed);

int BDEXCH(double AVDEPM,double W,double TAU,double TAUCD,double TAUCS,
		   double M,double VOL,double FRCSED,double *SUSP,double *BED,
		   double *DEPSCR);
int sandld(double isand,double vols,double srovol,double vol,double erovol,
		   double ksand,double avvele,double expsnd,double rom,int sandfg,
		   double db50e,double hrade,double slope,double tw, double wsande,
		   double twide,double db50m,double fsl,double avdepe,double *sand,
		   double *rsand,double *bdsand,double *depscr,double *rosand);
int toffal(double *v,double *fdiam,double *fhrad,double *slope,double *tempr, 
		   double *vset,double *gsi);
int colby(double *v,double *db50m,double *fhrad,double *fsl,double *tempr, 
		  double *gsi,int *ferror,int *d50err,int *hrerr,int *velerr);

