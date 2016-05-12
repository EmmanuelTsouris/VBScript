
Configure SMS Client Cache
Emmanuel Tsouris

Command Line Parameters
	CacheMin: Size to check for
	CacheSize: Size to set if Cache is less than CacheMin

Usage Examples
	configureclientcachesize.vbs /CacheMin:1024 /CacheSize:1024
	configureclientcachesize.vbs /CacheSize:1024

Summary

A client side script is necessary to reconfigure the client side SMS Cache. The cache is used for download & execute operations and must be adjusted accordinly when distributing large packages.
