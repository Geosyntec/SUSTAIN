//-----------------------------------------------------------------------------
//   findroot.h
//
//   Header file for root finding method contained in findroot.c
//-----------------------------------------------------------------------------

int findroot_Newton(float x1, float x2, float* rts, float xacc,
                    void (*func) (float x, float* f, float* df) );
float findroot_Ridder(float x1, float x2, float xacc, float (*func)(float));

