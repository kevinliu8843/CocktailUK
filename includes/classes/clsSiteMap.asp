<%
Class clsSiteMap
	Private m_strRoot, m_strMasterFile, m_strFile, m_strFileIndex, fso
	Private m_strIndex, m_strUrls, m_intNumUrls, m_strCurrentFile
	Private m_blnIndexOpen, m_blnFileOpen

	Public Sub Class_Initialize()
		Call Reset()
	End Sub
	
	Public Sub Reset()
		m_strRoot 		= Server.MapPath("/") & "\"
		m_strMasterFile	= ""
		m_strFile		= "sitemap"
		m_strCurrentFile= ""
		m_strFileIndex	= 0
		m_intNumUrls	= 0
		m_blnIndexOpen	= False
		m_blnFileOpen	= False
		m_strIndex		= ""
		Set fso			= Server.CreateObject("Scripting.FileSystemObject")
	End Sub
	
	Public Sub Class_Terminate()
		If m_blnIndexOpen Then
			Call CloseMapIndex()
		End If
		Set fso = Nothing
	End Sub
	
	Public Sub AddSiteMapURL(strURL, dteLastMod, strFrequency, strPriority)
		If NOT m_blnFileOpen Then
			Call NewMapIndex()
		End If
		
		m_intNumUrls = m_intNumUrls + 1
		
		strURL = Replace(strURL, "&", "&amp;")
		
		m_strUrls = m_strUrls & "<url><loc>" & strURL & "</loc>"
		If FormatDate(dteLastMod) <> "" Then
			m_strUrls = m_strUrls & "<lastmod>" & FormatDate(dteLastMod) & "</lastmod>"
		End If
		m_strUrls = m_strUrls & "<changefreq>" & strFrequency & "</changefreq><priority>" & strPriority & "</priority></url>" & VbCrLf
		
		If m_intNumUrls = 100 Then
			Call AppendToMapFile(m_strUrls)
			m_strUrls = ""
		End If
		
		If CurrentFileSize() > 9500000 Then
			Call CloseMapFile()
		End If
	End Sub
	
	Private Sub NewMapFile()
		m_strUrls = m_strUrls & "<?xml version=""1.0"" encoding=""UTF-8""?><urlset  xmlns=""http://www.sitemaps.org/schemas/sitemap/0.9"">" & VbCrLf
		Call SaveTextFile(m_strCurrentFile, m_strUrls)

		m_strUrls = ""
		m_blnFileOpen = True
	End Sub
	
	Private Sub CloseMapFile()
		m_strUrls = m_strUrls & "</urlset>"
		Call AppendToMapFile(m_strUrls)
		m_strCurrentFile = ""
		m_blnFileOpen = False
	End Sub
	
	Private Sub AppendToMapFile(strURLs)
		Call AppendTextFile(m_strCurrentFile, strURLs)
	End Sub
	
	Private Sub NewMapIndex()
		Dim strFileIndex
		
		If m_strIndex = "" Then
			m_strIndex = "<?xml version=""1.0"" encoding=""UTF-8""?><sitemapindex xmlns=""http://www.sitemaps.org/schemas/sitemap/0.9"">" & VbCrLf
		End If
		
		m_strFileIndex 		= m_strFileIndex + 1
		If m_strFileIndex = 1 Then
			strFileIndex = ""
		Else
			strFileIndex = m_strFileIndex
		End If
		
		m_strCurrentFile 	= m_strRoot & m_strFile & strFileIndex & ".xml"
		m_strIndex 			= m_strIndex & "<sitemap><loc>" & strUnSecureURL & "sitemap" & strFileIndex & ".xml</loc><lastmod>" & FormatDate(Now()) & "</lastmod></sitemap>" & VbCrLf
		
		Call NewMapFile()
		
		m_blnIndexOpen = True
	End Sub
	
	Private Sub CloseMapIndex()
		If m_strFileIndex > 1 Then
			m_strIndex = m_strIndex & "</sitemapindex>"
			Call SaveTextFile(m_strRoot & "sitemap_index.xml", m_strIndex)
		End If
		
		If m_blnFileOpen Then
			Call CloseMapFile()
		End If
		
		m_blnIndexOpen = False
	End Sub
	
	Private Function CurrentFileSize()
		Dim objFile
		CurrentFileSize = 0
		
		If m_blnFileOpen AND fso.FileExists(m_strCurrentFile) Then
			Set objFile = fso.GetFile(m_strCurrentFile)
			CurrentFileSize = objFile.Size
			Set objFile = Nothing
		End If
	End Function
	
	Private Function FormatDate(dteNow)
		If IsDate(dteNow) Then
			FormatDate = Year(dteNow) & "-"
			
			If Month(dteNow) <= 9 Then
				FormatDate = FormatDate & "0" & Month(dteNow) & "-"
			Else
				FormatDate = FormatDate & Month(dteNow) & "-"
			End If
			
			If Day(dteNow) <= 9 Then
					FormatDate = FormatDate & "0" & Day(dteNow)
			Else
				FormatDate = FormatDate & Day(dteNow)
			End If
		Else
			FormatDate = ""
		End If
	End Function
End Class
%>