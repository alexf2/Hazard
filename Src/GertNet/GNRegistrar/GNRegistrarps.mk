
GNRegistrarps.dll: dlldata.obj GNRegistrar_p.obj GNRegistrar_i.obj
	link /dll /out:GNRegistrarps.dll /def:GNRegistrarps.def /entry:DllMain dlldata.obj GNRegistrar_p.obj GNRegistrar_i.obj \
		kernel32.lib rpcndr.lib rpcns4.lib rpcrt4.lib oleaut32.lib uuid.lib \

.c.obj:
	cl /c /Ox /DWIN32 /D_WIN32_WINNT=0x0400 /DREGISTER_PROXY_DLL \
		$<

clean:
	@del GNRegistrarps.dll
	@del GNRegistrarps.lib
	@del GNRegistrarps.exp
	@del dlldata.obj
	@del GNRegistrar_p.obj
	@del GNRegistrar_i.obj
