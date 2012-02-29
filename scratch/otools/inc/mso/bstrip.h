#ifndef BSTRIP_DEFINED
#define BSTRIP_DEFINED
/****************************************************************************
    HBSTRIP is an opaque reference to an Office Bitmap record.
 ****************************************************************************/
struct BSTRIP
{
   int olb;          // bitmap implementation flags
   int cbmp;         // number of sub-bitmaps in the hbmp
   int dx, dy;       // size of the bitmap
   int dxSub, dySub; // Size of each subbitmap	
   void* psm;        // Pointer to shared memory
};

typedef struct BSTRIP* HBSTRIP;
#define hbstripNil ((HBSTRIP)NULL)
#endif
