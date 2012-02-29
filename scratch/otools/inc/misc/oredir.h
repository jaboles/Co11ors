#define MAKEWZ2(X)      L##X
#define MAKEWZ(X)       MAKEWZ2(X)

#define MAKESZ2(X)      #X
#define MAKESZ(X)       MAKESZ2(X)

#define OREDIR_FQDN     r.officeint.microsoft.com
#define SZOREDIR_FQDN   MAKESZ(OREDIR_FQDN)
#define WZOREDIR_FQDN   MAKEWZ(SZOREDIR_FQDN)
