<%
class aspZip
	dim BlankZip, NoInterfaceYesToAll
	dim fso, curArquieve, created, saved
	dim files, m_path, zipApp, zipFile	
	
	public property get Count()
		Count = files.Count
	end property
	
	public property get Path
		Path = m_path
	end property
	
	private sub class_initialize()
		BlankZip = Chr(80) & Chr(75) & Chr(5) & Chr(6) & String(18, 0) 	
		NoInterfaceYesToAll = 4 or 16 or 1024
		set fso = createObject("scripting.filesystemobject")
		set files = createObject("Scripting.Dictionary")
		Set zipApp = CreateObject("Shell.Application")
	end sub
	
	private sub class_terminate()
		set curArquieve = nothing
		set zipApp = nothing
		set files = nothing
		if created and not saved then
			on error resume next
			fso.deleteFile m_path
			on error goto 0
		end if
		set fso = nothing
	end sub
	
	public sub OpenArquieve(byval path)
		dim file
		path = replace(path, "/", "\")
		m_path = Server.MapPath(path)
		if not fso.fileexists(m_path) then
			set file = fso.createTextFile(m_path)
			file.write BlankZip
			file.close()
			set file = nothing
			
			set curArquieve = zipApp.NameSpace(m_path)
			created = true
		else
			dim cnt
			set curArquieve = zipApp.NameSpace(m_path)
			cnt = 0
			for each file in curArquieve.Items
				cnt = cnt + 1
				files.add file.path, cnt
			next
		end if
		saved = false
	end sub
	
	public sub Add(byval path)
		path = replace(path, "/", "\")		
		if instr(path, ":") = 0 then path = Server.mappath(path)
		
		if not fso.fileExists(path) and not fso.folderExists(path) then
			err.raise 1, "File not exists", "The input file name doen't correspond to an existing file"
			
		elseif not files.exists(path) Then
			files.add path, files.Count + 1
		end if
	end sub
	
	public sub Remove(byval path)
		if files.exists(path) then files.Remove(path)
	end sub
	
	public sub RemoveAll()
		files.RemoveAll()
	end sub
	
	public sub CloseArquieve()
		dim filepath, file, initTime, fileCount
		dim cnt
		cnt = 0
		For Each filepath In files.keys
			if instr(filepath, m_path) = 0 then
				curArquieve.Copyhere filepath, NoInterfaceYesToAll
				fileCount = curArquieve.items.Count
				On Error Resume Next
					wscript.sleep(10)
					cn = cnt + 1
				On Error GoTo 0
			end if
		next
		
		saved = true
	end sub
		
	public sub ExtractTo(byval path)
		if typeName(curArquieve) = "Folder3" Then
			path = Server.MapPath(path)
			if not fso.folderExists(path) then
				fso.createFolder(path)
			end if
			zipApp.NameSpace(path).CopyHere curArquieve.Items, NoInterfaceYesToAll
		end if
	end sub
end class

dim zip, backfilepath
backfilepath = "webBackup.zip" ' 打包文件
set zip = new aspZip
zip.OpenArquieve(backfilepath)
zip.Add("..\webroot") ' 需要打包的目录或文件,可增加多行
' 如再加一行,会向zip追加添加的文件
' zip.Add("web.config")
zip.CloseArquieve()
set zip = nothing
response.write "ok"
%>
