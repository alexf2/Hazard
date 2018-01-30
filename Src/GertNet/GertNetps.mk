
GertNetps.dll: dlldata.obj GertNet_p.obj GertNet_i.obj
	link /dll /out:GertNetps.dll /def:GertNetps.def /entry:DllMain dlldata.obj GertNet_p.obj GertNet_i.obj \
		kernel32.lib rpcndr.lib rpcns4.lib rpcrt4.lib oleaut32.lib uuid.lib \

.c.obj:
	cl /c /Ox /DWIN32 /D_WIN32_WINNT=0x0400 /DREGISTER_PROXY_DLL \
		$<

clean:
	@del GertNetps.dll
	@del GertNetps.lib
	@del GertNetps.exp
	@del dlldata.obj
	@del GertNet_p.obj
	@del GertNet_i.obj
