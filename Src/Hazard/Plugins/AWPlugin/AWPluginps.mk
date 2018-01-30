
AWPluginps.dll: dlldata.obj AWPlugin_p.obj AWPlugin_i.obj
	link /dll /out:AWPluginps.dll /def:AWPluginps.def /entry:DllMain dlldata.obj AWPlugin_p.obj AWPlugin_i.obj \
		kernel32.lib rpcndr.lib rpcns4.lib rpcrt4.lib oleaut32.lib uuid.lib \

.c.obj:
	cl /c /Ox /DWIN32 /D_WIN32_WINNT=0x0400 /DREGISTER_PROXY_DLL \
		$<

clean:
	@del AWPluginps.dll
	@del AWPluginps.lib
	@del AWPluginps.exp
	@del dlldata.obj
	@del AWPlugin_p.obj
	@del AWPlugin_i.obj
