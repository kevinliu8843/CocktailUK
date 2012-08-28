<%
Sub writeCOW(l_rs, l_cn, wkNum)
	Dim intCOWID, strName, sql

	intCOWID = getCOWID(wkNum)

	sql = "SELECT name, ID, type from cocktail WHERE ID=" & intCOWID
	l_rs.Open sql, l_cn, 0, 3
	strName = Replace( replaceStuffBack( l_rs("name") ), ",", "" )
	Response.Write "<A HREF=""/db/viewCocktail.asp?ID=" & l_rs("ID") & """>" & capitalise(strName) & "</A>"
	l_rs.Close
End Sub
	
Sub writeImgCOW(l_rs, l_cn, wkNum)
	Dim intCOWID, strName, FSO, strImgSrc, sql

	intCOWID = getCOWID(wkNum)
	sql = "SELECT name, ID, type from cocktail WHERE ID=" & intCOWID
	l_rs.Open sql, l_cn, 0, 3
	Set FSO = Server.CreateObject ("Scripting.filesystemobject")
	strName = Replace( replaceStuffBack( l_rs("name") ), ",", "" )

	If FSO.FileExists(Server.mappath("/images/cocktailThumbs/" & strName & ".jpg" ) ) THEN
		strImgSrc = "/images/cocktailThumbs/" & strName & ".jpg"
	Else
		If l_rs("type") = 1 THEN	
			strImgSrc = "/images/cocktailThumbs/cocktail.jpg"
		Else
			strImgSrc = "/images/cocktailThumbs/shooter.jpg"
		End If
	End If
	
	Response.Write "<A HREF=""/db/viewCocktail.asp?ID="&l_rs("ID")&"""><IMG border=1 SRC=""" & Replace(strImgSrc, " ", "%20") & """ alt=""Cocktail of the day : " & capitalise(strName) & """></A>"
	l_rs.Close
	Set FSO = Nothing
End Sub

Function getCOWID(wkNum)
	Select case wkNum
		case 1 getCOWID=14 ' 01/01/2002
		case 2 getCOWID=270 ' 02/01/2002
		case 3 getCOWID=209 ' 03/01/2002
		case 4 getCOWID=40 ' 04/01/2002
		case 5 getCOWID=267 ' 05/01/2002
		case 6 getCOWID=218 ' 06/01/2002
		case 7 getCOWID=13 ' 07/01/2002
		case 8 getCOWID=69 ' 08/01/2002
		case 9 getCOWID=255 ' 09/01/2002
		case 10 getCOWID=253 ' 10/01/2002
		case 11 getCOWID=185 ' 11/01/2002
		case 12 getCOWID=111 ' 12/01/2002
		case 13 getCOWID=257 ' 13/01/2002
		case 14 getCOWID=98 ' 14/01/2002
		case 15 getCOWID=166 ' 15/01/2002
		case 16 getCOWID=243 ' 16/01/2002
		case 17 getCOWID=143 ' 17/01/2002
		case 18 getCOWID=237 ' 18/01/2002
		case 19 getCOWID=161 ' 19/01/2002
		case 20 getCOWID=183 ' 20/01/2002
		case 21 getCOWID=43 ' 21/01/2002
		case 22 getCOWID=258 ' 22/01/2002
		case 23 getCOWID=244 ' 23/01/2002
		case 24 getCOWID=212 ' 24/01/2002
		case 25 getCOWID=31 ' 25/01/2002
		case 26 getCOWID=15 ' 26/01/2002
		case 27 getCOWID=137 ' 27/01/2002
		case 28 getCOWID=254 ' 28/01/2002
		case 29 getCOWID=107 ' 29/01/2002
		case 30 getCOWID=139 ' 30/01/2002
		case 31 getCOWID=118 ' 31/01/2002
		case 32 getCOWID=181 ' 01/02/2002
		case 33 getCOWID=198 ' 02/02/2002
		case 34 getCOWID=277 ' 03/02/2002
		case 35 getCOWID=113 ' 04/02/2002
		case 36 getCOWID=121 ' 05/02/2002
		case 37 getCOWID=172 ' 06/02/2002
		case 38 getCOWID=225 ' 07/02/2002
		case 39 getCOWID=197 ' 08/02/2002
		case 40 getCOWID=76 ' 09/02/2002
		case 41 getCOWID=160 ' 10/02/2002
		case 42 getCOWID=260 ' 11/02/2002
		case 43 getCOWID=89 ' 12/02/2002
		case 44 getCOWID=177 ' 13/02/2002
		case 45 getCOWID=139 ' 14/02/2002
		case 46 getCOWID=105 ' 15/02/2002
		case 47 getCOWID=275 ' 16/02/2002
		case 48 getCOWID=45 ' 17/02/2002
		case 49 getCOWID=48 ' 18/02/2002
		case 50 getCOWID=39 ' 19/02/2002
		case 51 getCOWID=66 ' 20/02/2002
		case 52 getCOWID=264 ' 21/02/2002
		case 53 getCOWID=227 ' 22/02/2002
		case 54 getCOWID=265 ' 23/02/2002
		case 55 getCOWID=149 ' 24/02/2002
		case 56 getCOWID=219 ' 25/02/2002
		case 57 getCOWID=215 ' 26/02/2002
		case 58 getCOWID=22 ' 27/02/2002
		case 59 getCOWID=42 ' 28/02/2002
		case 60 getCOWID=57 ' 01/03/2002
		case 61 getCOWID=12 ' 02/03/2002
		case 62 getCOWID=146 ' 03/03/2002
		case 63 getCOWID=228 ' 04/03/2002
		case 64 getCOWID=188 ' 05/03/2002
		case 65 getCOWID=102 ' 06/03/2002
		case 66 getCOWID=5 ' 07/03/2002
		case 67 getCOWID=18 ' 08/03/2002
		case 68 getCOWID=128 ' 09/03/2002
		case 69 getCOWID=2 ' 10/03/2002
		case 70 getCOWID=27 ' 11/03/2002
		case 71 getCOWID=101 ' 12/03/2002
		case 72 getCOWID=157 ' 13/03/2002
		case 73 getCOWID=269 ' 14/03/2002
		case 74 getCOWID=63 ' 15/03/2002
		case 75 getCOWID=273 ' 16/03/2002
		case 76 getCOWID=200 ' 17/03/2002
		case 77 getCOWID=67 ' 18/03/2002
		case 78 getCOWID=187 ' 19/03/2002
		case 79 getCOWID=180 ' 20/03/2002
		case 80 getCOWID=52 ' 21/03/2002
		case 81 getCOWID=231 ' 22/03/2002
		case 82 getCOWID=46 ' 23/03/2002
		case 83 getCOWID=249 ' 24/03/2002
		case 84 getCOWID=271 ' 25/03/2002
		case 85 getCOWID=131 ' 26/03/2002
		case 86 getCOWID=272 ' 27/03/2002
		case 87 getCOWID=54 ' 28/03/2002
		case 88 getCOWID=21 ' 29/03/2002
		case 89 getCOWID=119 ' 30/03/2002
		case 90 getCOWID=103 ' 31/03/2002
		case 91 getCOWID=225 ' 01/04/2002
		case 92 getCOWID=64 ' 02/04/2002
		case 93 getCOWID=196 ' 03/04/2002
		case 94 getCOWID=153 ' 04/04/2002
		case 95 getCOWID=132 ' 05/04/2002
		case 96 getCOWID=195 ' 06/04/2002
		case 97 getCOWID=7 ' 07/04/2002
		case 98 getCOWID=86 ' 08/04/2002
		case 99 getCOWID=202 ' 09/04/2002
		case 100 getCOWID=135 ' 10/04/2002
		case 101 getCOWID=186 ' 11/04/2002
		case 102 getCOWID=34 ' 12/04/2002
		case 103 getCOWID=6 ' 13/04/2002
		case 104 getCOWID=164 ' 14/04/2002
		case 105 getCOWID=174 ' 15/04/2002
		case 106 getCOWID=70 ' 16/04/2002
		case 107 getCOWID=178 ' 17/04/2002
		case 108 getCOWID=207 ' 18/04/2002
		case 109 getCOWID=250 ' 19/04/2002
		case 110 getCOWID=194 ' 20/04/2002
		case 111 getCOWID=85 ' 21/04/2002
		case 112 getCOWID=133 ' 22/04/2002
		case 113 getCOWID=136 ' 23/04/2002
		case 114 getCOWID=127 ' 24/04/2002
		case 115 getCOWID=154 ' 25/04/2002
		case 116 getCOWID=278 ' 26/04/2002
		case 117 getCOWID=36 ' 27/04/2002
		case 118 getCOWID=75 ' 28/04/2002
		case 119 getCOWID=237 ' 29/04/2002
		case 120 getCOWID=28 ' 30/04/2002
		case 121 getCOWID=24 ' 01/05/2002
		case 122 getCOWID=110 ' 02/05/2002
		case 123 getCOWID=130 ' 03/05/2002
		case 124 getCOWID=145 ' 04/05/2002
		case 125 getCOWID=100 ' 05/05/2002
		case 126 getCOWID=234 ' 06/05/2002
		case 127 getCOWID=37 ' 07/05/2002
		case 128 getCOWID=276 ' 08/05/2002
		case 129 getCOWID=191 ' 09/05/2002
		case 130 getCOWID=94 ' 10/05/2002
		case 131 getCOWID=106 ' 11/05/2002
		case 132 getCOWID=216 ' 12/05/2002
		case 133 getCOWID=91 ' 13/05/2002
		case 134 getCOWID=115 ' 14/05/2002
		case 135 getCOWID=190 ' 15/05/2002
		case 136 getCOWID=171 ' 16/05/2002
		case 137 getCOWID=79 ' 17/05/2002
		case 138 getCOWID=151 ' 18/05/2002
		case 139 getCOWID=8 ' 19/05/2002
		case 140 getCOWID=9 ' 20/05/2002
		case 141 getCOWID=155 ' 21/05/2002
		case 142 getCOWID=201 ' 22/05/2002
		case 143 getCOWID=268 ' 23/05/2002
		case 144 getCOWID=140 ' 24/05/2002
		case 145 getCOWID=246 ' 25/05/2002
		case 146 getCOWID=134 ' 26/05/2002
		case 147 getCOWID=263 ' 27/05/2002
		case 148 getCOWID=80 ' 28/05/2002
		case 149 getCOWID=219 ' 29/05/2002
		case 150 getCOWID=82 ' 30/05/2002
		case 151 getCOWID=142 ' 31/05/2002
		case 152 getCOWID=163 ' 01/06/2002
		case 153 getCOWID=163 ' 02/06/2002
		case 154 getCOWID=163 ' 03/06/2002
		case 155 getCOWID=163 ' 04/06/2002
		case 156 getCOWID=152 ' 05/06/2002
		case 157 getCOWID=210 ' 06/06/2002
		case 158 getCOWID=241 ' 07/06/2002
		case 159 getCOWID=221 ' 08/06/2002
		case 160 getCOWID=4 ' 09/06/2002
		case 161 getCOWID=95 ' 10/06/2002
		case 162 getCOWID=175 ' 11/06/2002
		case 163 getCOWID=216 ' 12/06/2002
		case 164 getCOWID=224 ' 13/06/2002
		case 165 getCOWID=274 ' 14/06/2002
		case 166 getCOWID=122 ' 15/06/2002
		case 167 getCOWID=94 ' 16/06/2002
		case 168 getCOWID=252 ' 17/06/2002
		case 169 getCOWID=262 ' 18/06/2002
		case 170 getCOWID=158 ' 19/06/2002
		case 171 getCOWID=192 ' 20/06/2002
		case 172 getCOWID=16 ' 21/06/2002
		case 173 getCOWID=60 ' 22/06/2002
		case 174 getCOWID=3 ' 23/06/2002
		case 175 getCOWID=173 ' 24/06/2002
		case 176 getCOWID=147 ' 25/06/2002
		case 177 getCOWID=150 ' 26/06/2002
		case 178 getCOWID=141 ' 27/06/2002
		case 179 getCOWID=168 ' 28/06/2002
		case 180 getCOWID=88 ' 29/06/2002
		case 181 getCOWID=124 ' 30/06/2002
		case 182 getCOWID=163 ' 01/07/2002
		case 183 getCOWID=251 ' 02/07/2002
		case 184 getCOWID=116 ' 03/07/2002
		case 185 getCOWID=112 ' 04/07/2002
		case 186 getCOWID=199 ' 05/07/2002
		case 187 getCOWID=144 ' 06/07/2002
		case 188 getCOWID=159 ' 07/07/2002
		case 189 getCOWID=114 ' 08/07/2002
		case 190 getCOWID=249 ' 09/07/2002
		case 191 getCOWID=125 ' 10/07/2002
		case 192 getCOWID=11 ' 11/07/2002
		case 193 getCOWID=279 ' 12/07/2002
		case 194 getCOWID=182 ' 13/07/2002
		case 195 getCOWID=120 ' 14/07/2002
		case 196 getCOWID=231 ' 15/07/2002
		case 197 getCOWID=179 ' 16/07/2002
		case 198 getCOWID=129 ' 17/07/2002
		case 199 getCOWID=204 ' 18/07/2002
		case 200 getCOWID=259 ' 19/07/2002
		case 201 getCOWID=167 ' 20/07/2002
		case 202 getCOWID=165 ' 21/07/2002
		case 203 getCOWID=97 ' 22/07/2002
		case 204 getCOWID=97 ' 23/07/2002
		case 205 getCOWID=243 ' 24/07/2002
		case 206 getCOWID=10 ' 25/07/2002
		case 207 getCOWID=208 ' 26/07/2002
		case 208 getCOWID=80 ' 27/07/2002
		case 209 getCOWID=186 ' 28/07/2002
		case 210 getCOWID=74 ' 29/07/2002
		case 211 getCOWID=203 ' 30/07/2002
		case 212 getCOWID=21 ' 31/07/2002
		case 213 getCOWID=85 ' 01/08/2002
		case 214 getCOWID=22 ' 02/08/2002
		case 215 getCOWID=8 ' 03/08/2002
		case 216 getCOWID=254 ' 04/08/2002
		case 217 getCOWID=73 ' 05/08/2002
		case 218 getCOWID=57 ' 06/08/2002
		case 219 getCOWID=180 ' 07/08/2002
		case 220 getCOWID=92 ' 08/08/2002
		case 221 getCOWID=150 ' 09/08/2002
		case 222 getCOWID=181 ' 10/08/2002
		case 223 getCOWID=161 ' 11/08/2002
		case 224 getCOWID=223 ' 12/08/2002
		case 225 getCOWID=35 ' 13/08/2002
		case 226 getCOWID=41 ' 14/08/2002
		case 227 getCOWID=156 ' 15/08/2002
		case 228 getCOWID=164 ' 16/08/2002
		case 229 getCOWID=214 ' 17/08/2002
		case 230 getCOWID=62 ' 18/08/2002
		case 231 getCOWID=239 ' 19/08/2002
		case 232 getCOWID=118 ' 20/08/2002
		case 233 getCOWID=202 ' 21/08/2002
		case 234 getCOWID=98 ' 22/08/2002
		case 235 getCOWID=132 ' 23/08/2002
		case 236 getCOWID=235 ' 24/08/2002
		case 237 getCOWID=279 ' 25/08/2002
		case 238 getCOWID=222 ' 26/08/2002
		case 239 getCOWID=113 ' 27/08/2002
		case 240 getCOWID=87 ' 28/08/2002
		case 241 getCOWID=90 ' 29/08/2002
		case 242 getCOWID=81 ' 30/08/2002
		case 243 getCOWID=108 ' 31/08/2002
		case 244 getCOWID=28 ' 01/09/2002
		case 245 getCOWID=269 ' 02/09/2002
		case 246 getCOWID=29 ' 03/09/2002
		case 247 getCOWID=192 ' 04/09/2002
		case 248 getCOWID=56 ' 05/09/2002
		case 249 getCOWID=257 ' 06/09/2002
		case 250 getCOWID=65 ' 07/09/2002
		case 251 getCOWID=84 ' 08/09/2002
		case 252 getCOWID=99 ' 09/09/2002
		case 253 getCOWID=54 ' 10/09/2002
		case 254 getCOWID=189 ' 11/09/2002
		case 255 getCOWID=65 ' 12/09/2002
		case 256 getCOWID=230 ' 13/09/2002
		case 257 getCOWID=145 ' 14/09/2002
		case 258 getCOWID=122 ' 15/09/2002
		case 259 getCOWID=60 ' 16/09/2002
		case 260 getCOWID=171 ' 17/09/2002
		case 261 getCOWID=119 ' 18/09/2002
		case 262 getCOWID=69 ' 19/09/2002
		case 263 getCOWID=144 ' 20/09/2002
		case 264 getCOWID=199 ' 21/09/2002
		case 265 getCOWID=107 ' 22/09/2002
		case 266 getCOWID=105 ' 23/09/2002
		case 267 getCOWID=37 ' 24/09/2002
		case 268 getCOWID=242 ' 25/09/2002
		case 269 getCOWID=109 ' 26/09/2002
		case 270 getCOWID=229 ' 27/09/2002
		case 271 getCOWID=18 ' 28/09/2002
		case 272 getCOWID=94 ' 29/09/2002
		case 273 getCOWID=274 ' 30/09/2002
		case 274 getCOWID=88 ' 01/10/2002
		case 275 getCOWID=13 ' 02/10/2002
		case 276 getCOWID=35 ' 03/10/2002
		case 277 getCOWID=174 ' 04/10/2002
		case 278 getCOWID=110 ' 05/10/2002
		case 279 getCOWID=96 ' 06/10/2002
		case 280 getCOWID=63 ' 07/10/2002
		case 281 getCOWID=162 ' 08/10/2002
		case 282 getCOWID=146 ' 09/10/2002
		case 283 getCOWID=268 ' 10/10/2002
		case 284 getCOWID=106 ' 11/10/2002
		case 285 getCOWID=238 ' 12/10/2002
		case 286 getCOWID=270 ' 13/10/2002
		case 287 getCOWID=249 ' 14/10/2002
		case 288 getCOWID=32 ' 15/10/2002
		case 289 getCOWID=50 ' 16/10/2002
		case 290 getCOWID=129 ' 17/10/2002
		case 291 getCOWID=244 ' 18/10/2002
		case 292 getCOWID=252 ' 19/10/2002
		case 293 getCOWID=24 ' 20/10/2002
		case 294 getCOWID=76 ' 21/10/2002
		case 295 getCOWID=49 ' 22/10/2002
		case 296 getCOWID=206 ' 23/10/2002
		case 297 getCOWID=12 ' 24/10/2002
		case 298 getCOWID=112 ' 25/10/2002
		case 299 getCOWID=220 ' 26/10/2002
		case 300 getCOWID=44 ' 27/10/2002
		case 301 getCOWID=14 ' 28/10/2002
		case 302 getCOWID=236 ' 29/10/2002
		case 303 getCOWID=127 ' 30/10/2002
		case 304 getCOWID=176 ' 31/10/2002
		case 305 getCOWID=179 ' 01/11/2002
		case 306 getCOWID=170 ' 02/11/2002
		case 307 getCOWID=196 ' 03/11/2002
		case 308 getCOWID=116 ' 04/11/2002
		case 309 getCOWID=78 ' 05/11/2002
		case 310 getCOWID=117 ' 06/11/2002
		case 311 getCOWID=1 ' 07/11/2002
		case 312 getCOWID=71 ' 08/11/2002
		case 313 getCOWID=66 ' 09/11/2002
		case 314 getCOWID=153 ' 10/11/2002
		case 315 getCOWID=173 ' 11/11/2002
		case 316 getCOWID=187 ' 12/11/2002
		case 317 getCOWID=143 ' 13/11/2002
		case 318 getCOWID=277 ' 14/11/2002
		case 319 getCOWID=79 ' 15/11/2002
		case 320 getCOWID=40 ' 16/11/2002
		case 321 getCOWID=233 ' 17/11/2002
		case 322 getCOWID=136 ' 18/11/2002
		case 323 getCOWID=149 ' 19/11/2002
		case 324 getCOWID=259 ' 20/11/2002
		case 325 getCOWID=133 ' 21/11/2002
		case 326 getCOWID=158 ' 22/11/2002
		case 327 getCOWID=232 ' 23/11/2002
		case 328 getCOWID=9 ' 24/11/2002
		case 329 getCOWID=121 ' 25/11/2002
		case 330 getCOWID=193 ' 26/11/2002
		case 331 getCOWID=125 ' 27/11/2002
		case 332 getCOWID=51 ' 28/11/2002
		case 333 getCOWID=198 ' 29/11/2002
		case 334 getCOWID=38 ' 30/11/2002
		case 335 getCOWID=32 ' 01/12/2002
		case 336 getCOWID=183 ' 02/12/2002
		case 337 getCOWID=83 ' 03/12/2002
		case 338 getCOWID=177 ' 04/12/2002
		case 339 getCOWID=101 ' 05/12/2002
		case 340 getCOWID=123 ' 06/12/2002
		case 341 getCOWID=262 ' 07/12/2002
		case 342 getCOWID=124 ' 08/12/2002
		case 343 getCOWID=184 ' 09/12/2002
		case 344 getCOWID=152 ' 10/12/2002
		case 345 getCOWID=250 ' 11/12/2002
		case 346 getCOWID=234 ' 12/12/2002
		case 347 getCOWID=77 ' 13/12/2002
		case 348 getCOWID=195 ' 14/12/2002
		case 349 getCOWID=47 ' 15/12/2002
		case 350 getCOWID=5 ' 16/12/2002
		case 351 getCOWID=263 ' 17/12/2002
		case 352 getCOWID=47 ' 18/12/2002
		case 353 getCOWID=138 ' 19/12/2002
		case 354 getCOWID=217 ' 20/12/2002
		case 355 getCOWID=53 ' 21/12/2002
		case 356 getCOWID=266 ' 22/12/2002
		case 357 getCOWID=38 ' 23/12/2002
		case 358 getCOWID=165 ' 24/12/2002
		case 359 getCOWID=137 ' 25/12/2002
		case 360 getCOWID=16 ' 26/12/2002
		case 361 getCOWID=26 ' 27/12/2002
		case 362 getCOWID=201 ' 28/12/2002
		case 363 getCOWID=30 ' 29/12/2002
		case 364 getCOWID=59 ' 30/12/2002
		case else getCOWID = 1	
	End Select
End Function
%>