
AnglFontps.dll: dlldata.obj AnglFont_p.obj AnglFont_i.obj
	link /dll /out:AnglFontps.dll /def:AnglFontps.def /entry:DllMain dlldata.obj AnglFont_p.obj AnglFont_i.obj \
		kernel32.lib rpcndr.lib rpcns4.lib rpcrt4.lib oleaut32.lib uuid.lib \

.c.obj:
	cl /c /Ox /DWIN32 /D_WIN32_WINNT=0x0400 /DREGISTER_PROXY_DLL \
		$<

clean:
	@del AnglFontps.dll
	@del AnglFontps.lib
	@del AnglFontps.exp
	@del dlldata.obj
	@del AnglFont_p.obj
	@del AnglFont_i.obj
