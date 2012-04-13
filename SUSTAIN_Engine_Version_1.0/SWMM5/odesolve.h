//-----------------------------------------------------------------------------
//  odesolve.h
//
//  Header file for ODE solver contained in odesolve.c
//
//-----------------------------------------------------------------------------

// functions that open, close, and use the ODE solver
int  odesolve_open(int n);
void odesolve_close(void);
int  odesolve_integrate(float ystart[], int n, float x1, float x2,
        float eps, float h1, void (*derivs)(float, float*, float*));
