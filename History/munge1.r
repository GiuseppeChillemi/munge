REBOL [
	Title:		"Munge function"
	Owner:		"Ashley G Truter"
	Version:	1.0.9
	Date:		20-Apr-2015
	Purpose:	"Extract and manipulate tabular values in blocks, delimited files and SQL Server tables."
	Tested: {
		2.7.8.3.1	R2			rebol.com
		2.101.0.3.1	R3 Alpha	rebolsource.net
		3.0.91.3.3	R3 64-bit	atronixengineering.com
	}
	Usage: {
		;	General
		cols?		Number of columns in a delimited file
		fields?		Column names in a delimited file
		sheets?		Number of sheets in an XLS file
		;	Associative array
		index		Converts a block of values into key and value pairs
		lookup		Finds a value in the index and returns the value after it
		assign		Adds/updates/removes a key and its associated value in an index
		retrieve	Return row from block using index to lookup key
		;	Munge
		execute		Execute a SQL statement (SQL Server, Access or SQLite)
		load-dsv	Loads delimiter-separated values from a file
		munge		Load and / or manipulate a block of tabular (column and row) values
		read-pdf	Reads from a PDF file
		read-xls	Reads from an XLS file
		sqlcmd		Execute a SQL Server statement
		unzip		Uncompress a file into a folder of the same name (BETA)
		worksheet	Add a worksheet to the current workbook
	}
	Licence:	"MIT. Free for both commercial and non-commercial use."
	History: {
		1.0.0	Initial release
		1.0.1	Renamed load-block to load-dsv
				Added new sqlcmd func
				Merged Excel funcs into new worksheet func
		1.0.2	Added /count
		1.0.3	Fixed /unqiue and /count
				Added /sum
		1.0.4	Added clean-path to load-dsv and worksheet
				Added /where integer! support (i.e. RowID)
				Added lookup, index funcs
				Added /merge
		1.0.5	Minor speed improvements
				Fixed minor merge bug
				Fixed /part make block! bug
				Removed string! as a data and save option
				Added /order
				Added /headings to sqlcmd
		1.0.6	SQL Server 2012 fix
				Refactored load-dsv based on csv-tools.r
				Added /max and /min
				Added /having
				Fixed /save to handle empty? buffer
				load-dsv now handles xls variants (e.g. xlsx, xlsm, xml, ods)
				Fixed bug with part/where/unique
				Added /compact
				Added console null print protection prior to all calls
				Added read-pdf
				Added read-xls
		1.0.7	Compatibility patches
					to-error		does not work in R3
					remove-each		R3 returns integer
					select			R2 /skip returns block
					unique			/skip broken
				Minor changes to work with R3
					read (R3 returns a binary)
					delete/any (not supported in R3)
					find/any (not working in R3)
					read/lines/part (not working in R3)
					call/show (not required or supported in R3)
					call/shell (required in R3 for *.vbs)
					call %file (call form %file works in R3)
				Removed /unique
				Added column name support
				Added /headings
				Added /save none target to return lines
				Merged /having into /group
				worksheet changes
					Removed columns argument
					Removed /widths and /footer refinements
					Added spec argument
					Added support for date and auto cell types
		1.0.8	Replaced to-error with cause-error
				Replaced func with funct
				Added execute function
				Added MS Access support to execute
				Added SQLite support to execute
				Added /only
				Added spec none! support
				Added /save none! support
				Fixed /merge bug
				Fixed sqlcmd /headings/key bug
				Added cols? function
				Added rows? function
				Added sheets? function
				Fixed to work with R3 Alpha (rebolsource.net)
				Added load-dsv /blocks
				Fixed delete/where (missing implied all)
				Added unzip
		1.0.9	Added call compatibility function for R3 Alpha
				Added /all support to read-xls
				Added /part to load-dsv
				Re-factored VBS calls
				Added fields? function
				Added associative array support (index, lookup, assign)
				Added unique undex support (index/direct, retrieve)
		1.0.10	Added slice function
				Added /sheet refinement to munge function
				Added string! support for /sheet refinement in all supported functions
				Fixed load-dsv gives more meaningful error if file doesn't exists
	}
]

;
;	Compatibility patches
;

unless any ["native" = mold get 'call find mold get 'call "/output"] [	; R3 Alpha fix
	*call: :call
	call: funct [command /output out /error err /wait /shell] [
		command: trim reform [either shell ["cmd /C"][""] command either any [output error]["1> $out.txt 2> $err.txt"][""]]
		rc: either wait [*call/wait command] [*call command]
		if exists? %$out.txt [
			all [output insert clear out read/string %$out.txt]
			delete %$out.txt
		]
		if exists? %$err.txt [
			all [error insert clear err read/string %$err.txt]
			delete %$err.txt
		]
		rc
	]
]

if integer? remove-each i [][][		; R3 fix
	*remove-each: :remove-each
	remove-each: func ['word data body][also data *remove-each :word data body]
]

if block? select/skip [0 0] 0 2 [	; R2 fix
	*select: :select
	select: func [series value /skip size][
		either skip [
			all [value: *select/skip series value size first value]
		][*select series value]
	]
]

if "1a" = unique/skip "1a1b" 2 [	; R2/R3 fix
	*unique: :unique
	unique: funct [set /skip size][
		either skip [
			row: make block! size
			repeat i size [append row to word! append form 'c i]
			skip: unset!
			do compose/deep [
				remove-each [(row)] sort/skip/all copy set size [
					either skip = reduce [(row)][true][skip: reduce [(row)] false]
				]
			]
		][*unique set]
	]
]

context [

	;
	;	Private
	;

	R64?: 3 = first system/version
	R3?: any [R64? 100 < second system/version]

	XLS?: func [
		file [file!]
	][
		find [%.xls %.xml %.ods] copy/part suffix? file 4
	]

	base-path: join to-local-file system/script/path "\"

	excel-metrics: 0x0	; sheets x columns

	sheet?: func [
		sheet [integer! string! none!]
	] [
		either string? sheet [ajoin [{"} sheet {"}]] [any [sheet 1]]
	]

	call-excel-vbs: func [
		file [file!]
		sheet [integer! string! none!]
		cmd [string!]
	][
		any [exists? file cause-error 'access 'cannot-open file]
		write %$tmp.vbs ajoin [
			{set X=CreateObject("Excel.Application"):X.DisplayAlerts=False:set W=X.Workbooks.Open("} to-local-file clean-path file {"):}
			{n=(X.Worksheets.Count*1000)+X.ActiveWorkbook.Worksheets(} sheet? sheet {).UsedRange.Columns.Count:}
			cmd
			":X.Workbooks.Close:WScript.Quit n"
		]
		excel-metrics: form call/wait/shell "$tmp.vbs"
		delete %$tmp.vbs
		excel-metrics: as-pair to-integer copy/part excel-metrics subtract length? excel-metrics 3 to-integer skip tail excel-metrics -3
	]

	to-xml-string: func [
		string [string!]
	][
		foreach [char code][
			"<"	"&lt;"
			">"	"&gt;"
			{"}	"&quot;"
		][replace/all string char code]
	]

	;
	;	Public
	;

	set 'cols? func [
		"Number of columns in a delimited file."
		file [file!]
		/sheet "Excel worksheet number or name (default is 1)"
			number [integer! string!]
	][
		either XLS? file [
			to-integer second call-excel-vbs file number ""
		][
			;	/lines/part returns n chars in R3
			length? load-dsv first either R3? [read/lines file][read/direct/lines/part file 1]
		]
	]

	set 'fields? funct [
		"Column names in a delimited file."
		file [file!]
		/sheet "Excel worksheet number of name(default is 1)"
			number [integer! string!]
	][
		any [exists? file cause-error 'access 'cannot-open file]
		either xls? file [
			call?: "x"
			unless any [R3? empty? call?][call/show clear call?]
			write/string %$tmp.vbs ajoin [{
				Set WB=CreateObject("Excel.Application")
				WB.DisplayAlerts=0
				WB.Workbooks.open "} to-local-file clean-path file {",false,true
				Set WS=WB.ActiveWorkbook.Worksheets(} sheet? number {)
				Line=""
				For Col=1 to WS.UsedRange.Columns.Count
					Line=Line&WS.Cells(1,Col).Value&VBTab
				Next
				WScript.Echo(LEFT(Line,(LEN(Line)-1)))
				WB.Workbooks.Close
			}]
			call/wait/output "cmd /C cscript $tmp.vbs //NoLogo" stdout: make string! 1000
			delete %$tmp.vbs
		][
			stdout: first either R3? [read/lines file][read/direct/lines/part file 1]
		]
		also load-dsv stdout stdout: none
	]

	set 'sheets? func [
		"Number of sheets in an XLS file."
		file [file!]
	][
		to-integer first call-excel-vbs file 1 ""
	]

	set 'slice func ["Extract part of a block" block [block!] len "Row length" [integer!] start [pair!] end [pair!] /local output part] [
		output: copy []
		if len < first start [return output]
		either positive? (first start) + (first end) - (len + 1) [
			part: 1 + len - (first start)
		] [
			part: first end
		]
		loop second end [
			append output copy/part at block first start part
			block: skip block len
			if tail? block [break]
		]

		new-line/all/skip output true part
	]


	;
	;	Associative array
	;
	
	either zero? select to-map [0 1 1 0] 1 [
		;
		;	R3 Map!
		;

		set 'index funct [
			"Converts a block of values into key and value pairs."
			block [block!]
			width [integer!] "Size of each entry (the skip) with width minus one being the key"
			/direct "Key / Rowid pairs"
		][
			all [not direct width < 2 cause-error 'user 'message "Width must be greater than 1"]
			any [mod size: length? block width cause-error 'user 'message "Width not a multiple of length"]
			blk: make block! 2 * rows: size / width
			either direct [
				either width = 1 [
					repeat i rows [
						append blk reduce [pick block i i]
					]
				][
					repeat i rows [
						append blk reduce [form copy/part skip block i * width - width width i]
					]
				]
			][
				either width = 2 [blk: block][
					size: width - 1
					repeat i rows [
						append blk reduce [form copy/part skip block i * width - width size pick block i * width]
					]
				]
			]
			also to-map blk blk: none
		]

		set 'lookup func [
			"Finds a value in the index and returns the value after it."
			index [map!]
			key
		][
			select index either block? key [form key][key]
		]

		set 'assign func [
			"Adds/updates/removes a key and its associated value in an index."
			index [map!]
			key
			value
		][
			append index reduce [either block? key [form key][key] value]
			value
		]

		set 'retrieve funct [
			"Return row from block using index to lookup key."
			block [block!]
			size [integer!]
			index [map!]
			key
		][
			all [block? key key: form key]
			all [idx: select index key copy/part skip block idx * size - size size]
		]
	][
		;
		;	R2 Hash!
		;

		set 'index funct [
			"Converts a block of values into key and value pairs."
			block [block!]
			width [integer!] "Size of each entry (the skip) with width minus one being the key"
			/direct "Key / Rowid pairs"
		][
			all [width = 1 return to-map block]
			any [mod size: length? block width cause-error 'user 'message "Width not a multiple of length"]
			rows: size / width
			either direct [
				blk: make block! rows
				repeat i rows [append blk form copy/part skip block i * width - width width]
				also to-map blk blk: none
			][
				blk: make map! rows * 2
				either width = 2 [
					foreach [key value] block [
						assign blk key value
					]
				][
					size: width - 1
					repeat i rows [
						assign blk copy/part skip block i * width - width size pick block i * width
					]
				]
				also blk blk: none
			]
		]

		set 'lookup func [
			"Finds a value in the index and returns the value after it."
			index [map!]
			key
		][
			select index as-binary uppercase form key
		]

		set 'assign funct [
			"Adds/updates/removes a key and its associated value in an index."
			index [map!]
			key
			value
		][
			idx: find index key: as-binary uppercase form key
			either value [
				either idx [poke index 1 + index? idx value][append index reduce [key value]]
			][all [idx remove/part idx 2]]
			value
		]

		set 'retrieve funct [
			"Return row from block using index to lookup key."
			block [block!]
			size [integer!]
			index [map!]
			key
		][
			all [block? key key: form key]
			all [idx: find index key copy/part skip block (index? idx) * size - size size]
		]
	]

	;
	;	Munge
	;

	set 'execute funct [
		"Execute a SQL statement (SQL Server, Access or SQLite)."
		database [string! file!]
		statement [string!]
		/key "Columns to convert to integer"
			columns [integer! block!]
		/headings "Keep column headings"
		/raw "Do not process return buffer"
	][
		call?: "x"
		unless any [R3? empty? call?][call/show clear call?]
		stderr: make string! 256
		case [
			string? database [			; SQL Server
				call/output/error reform ["sqlcmd -S" first parse database "/" "-d" second parse database "/" "-I -Q" ajoin [{"} trim/lines statement {"}] {-W -w 65535 -s"^-"}] buffer: make string! 8192 stderr
				all [empty? stderr "Msg" = copy/part buffer 3 stderr: trim/lines buffer]
				all [R3? trim/with buffer "^M"]
				replace/all buffer "^/^/" "^/" ; sqlcmd 2012 inserts blank lines every 4k or so rows
				replace/all buffer "^/^/" "^/"
			]
			%.accdb = suffix? database [; MS Access
				any [exists? database cause-error 'access 'cannot-open database]
				write %$tmp.vbs ajoin [{set A=CreateObject("ADODB.Connection"):A.Open "Provider=Microsoft.Ace.OLEDB.12.0;Data Source=} to-local-file clean-path database {":A.Execute "} copy/part statement index? find statement " from " {into [text;database=} to-local-file path: first split-path clean-path database {].[$tmp.txt]} find statement " from " {"}]
				write schema: join path %Schema.ini ajoin ["[$tmp.txt]^/ColNameHeader=" either headings ["True"]["False"] "^/Format=TabDelimited"]
				also call/wait/shell either R64? ["$tmp.vbs"]["%windir%\sysnative\wscript $tmp.vbs"] stderr delete %$tmp.vbs
				delete schema
				either exists? file: join path %$tmp.txt [
					buffer: read/string file
					delete file
				][cause-error 'access 'cannot-open file]
			]
			true [						; SQLite
				call/output/error reform compose [{sqlite3 -separator "^-"} (either headings ["-header"][]) to-local-file database ajoin [{"} trim/lines statement {"}]] buffer: make string! 8192 stderr
				all [R3? trim/with buffer "^M"]
			]
		]
		any [empty? stderr cause-error 'user 'message stderr]
		all [raw return buffer]
		either "select" = copy/part statement 6 [
			all ["^/" = buffer return make block! 0] ; SQLite empty result set
			row: copy/part buffer find buffer "^/"
			all [#"^-" = last row append row "^-"]
			size: length? parse/all row "^-"
			buffer: parse/all buffer "^-^/"
			;	clean SQLCMD buffer
			if string? database [
				either headings [remove/part skip buffer size size][remove/part buffer size * 2]
				remove back tail buffer
				foreach val buffer [all [val == "NULL" clear val]]
			]
			;	remove leading and trailing spaces
			foreach val buffer [
				trim val
			]
			;	replace string! with integer!
			if key [
				rows: divide length? buffer size
				all [headings -- rows]
				foreach offset to-block columns [
					all [headings offset: offset + size]
					loop rows [
						poke buffer offset to integer! pick buffer offset
						offset: offset + size
					]
				]
			]
			also buffer buffer: none
		][trim/with buffer "()^/"]
	]

	set 'load-dsv funct [
		[catch]
		"Loads delimiter-separated values from a file."
		file [file! string!]
		/delimit {Alternate delimiter (default is tab then comma)}
			delimiter [char!]
		/sheet "Excel worksheet (default is 1)"
			number [integer! string! none!]
		/part "Consecutive offset position(s) to retrieve"
			columns [block!]
		/blocks "Rows as blocks"
	][
		if file? file [
			unless exists? file [throw-error 'access 'cannot-open file]
			all [zero? size? file return make block! 0]
			file: either XLS? file [delimiter: #"," read-xls/sheet file number][read/string file]
		]
		all [empty? file return make block! 0]
		delimiter: form any [
			delimiter
			either find/part file "^-" any [find file "^/" file] [#"^-"][#","]
		]
		; Parse rules
		valchars: remove/part charset [#"^(00)" - #"^(FF)"] crlf
		valchars: compose [any (remove/part valchars delimiter)]
		value: either part [
			[
				; Value in quotes, with Excel-compatible handling of bad syntax
				{"} (clear val) x: [to {"} | to end] y: (insert/part tail val x y)
				any [{"} x: {"} [to {"} | to end] y: (insert/part tail val x y)]
				[{"} x: valchars y: (insert/part tail val x y) | end]
				(insert tail blk trim copy val) |
				; Raw value
				x: valchars y: (all [find columns ++ pos insert tail blk trim copy/part x y])
			]
		][
			[
				{"} (clear val) x: [to {"} | to end] y: (insert/part tail val x y)
				any [{"} x: {"} [to {"} | to end] y: (insert/part tail val x y)]
				[{"} x: valchars y: (insert/part tail val x y) | end]
				(insert tail blk trim copy val) |
				x: valchars y: (insert tail blk trim copy/part x y)
			]
		]
		val: make string! 1000
		pos: 1
		either blocks [
			output: make block! 1
			parse/all file [z: any [end break | (blk: copy []) value any [delimiter value][crlf | cr | lf | end] (pos: 1 output: insert/only output blk)] z:]
			blk: head output
		][
			blk: make block! 100000
			parse/all file [z: any [end break | value any [delimiter value][crlf | cr | lf | end] (pos: 1) ] z:]
		]
		also blk (file: output: blk: value: val: x: y: z: valchars: none)
	]

	set 'munge funct [
		"Load and/or manipulate a block of tabular (column and row) values."
		data [block! file!] "REBOL block, CSV or Excel file"
		spec [integer! block! none!] "Size of each record or block of heading words (none! gets cols? file)"
		/update "Offset/value pairs (returns original block)"
			action [block!]
		/delete "Delete matching rows (returns original block)"
		/part "Offset position(s) to retrieve"
			columns [block! integer! word!]
		/where "Rowid or expression that can reference columns as c1 or A, c2 or B, etc"
			condition [block! integer!]
		/headings "Returns heading words as first row (unless condition is integer)"
		/compact "Remove blank rows"
		/only "Remove duplicate rows"
		/merge "Join outer block (data) to inner block on keys"
			inner-block [block! map!] "Block to lookup values in"
			inner-size [integer!] "Size of each record"
			cols [block!] "Offset position(s) to retrieve in merged block"
			keys [block!] "Outer/inner join column pairs"
		/group "One of count, flip, max, min or sum"
			having [word! block!] "Word or expression that can reference the initial result set column as count, flip, max, etc"
		/order "Sort result set"
		/save "Write result to a delimited or Excel file"
			file [file! none!] "csv, xml, xlsx or tab delimited"
		/list "Return new-line records"
		/sheet "Excel worksheet (default is 1)"
			number [integer! string!]
	][
		if file? data [
			any [spec XLS? data spec: cols? data]
			data: either sheet [
				load-dsv/sheet data number
			] [
				load-dsv data
			]
			any [spec spec: to-integer second excel-metrics]
		]

		all [
			empty? data
			either save [write file "" exit][return data]
		]

		either integer? spec [
			spec: make block! size: spec
			repeat i size [append spec to word! append form 'c i]
		][size: length? spec]

		row: spec

		any [
			integer? rows: divide length? data size
			cause-error 'user 'message "size not a multiple of length"
		]

		all [
			not integer? columns
			repeat i length? columns: to block! columns [
				all [
					word? val: pick columns i
					3 > length? val: form val
					poke columns i either 1 = length? uppercase val [subtract to integer! first val 64][(26 * subtract to integer! first val 64) + subtract to integer! second val 64]
				]
			]
		]

		blk: copy []

		all [
			headings
			not update
			either part [foreach i to block! columns [append blk pick spec i]][repeat i size [append blk pick spec i]]
		]

		case [
			integer? condition [
				i: condition * size - size
				return case [
					update [
						foreach [col val] reduce action [
							poke data i + col val
						]
						data
					]
					delete [head remove/part skip data i size]
					part [
						either integer? columns [pick data i + columns][
							blk: make block! length? columns
							foreach col columns [
								append blk pick data i + col
							]
							blk
						]
					]
					true [copy/part skip data i size]
				]
			]
			update [
				foreach [col val] reduce action [
					append blk compose [
						poke data i + (col) (either datatype? val [compose [to (val) pick data i + (col)]][val])
					]
				]
				i: 0
				either where [
					do compose/deep [
						foreach [(row)] data [
							all [(condition) (blk)]
							i: i + (size)
						]
					]
				][loop rows compose [(blk) i: i + (size)]]
				return data
			]
			delete [
				either where [
					do compose/deep [
						remove-each [(row)] data [all [(condition)]]
					]
				][clear data]
				return data
			]
			part [
				either where [
					either block? columns [
						part: reduce ['reduce copy []]
						foreach col columns [
							append last part pick row col
						]
					][part: pick row columns]
					do compose/deep [
						foreach [(row)] data [
							all [
								(condition)
								append blk (part)
							]
						]
					]
				][
					either block? columns [
						part: reduce ['reduce copy []]
						foreach col columns [
							append last part either col = 'rowid [[rowid]][compose [pick data (col)]]
						]
					][part: compose [pick data (columns)]]
					repeat rowid rows compose [
						append blk (part)
						data: skip data (size)
					]
					data: head data
				]
				row: make block! size: either integer? columns [1][length? columns]
				foreach col to block! columns [append row either integer? col [pick spec col][col]]
				data: blk
			]
			where [
				do compose/deep [
					foreach [(row)] data [
						all [(condition) append blk reduce [(row)]]
					]
				]
				data: blk
			]
			headings [
				append blk data
				data: blk
			]
		]

		all [
			compact
			do compose/deep [remove-each [(row)] data [empty? form reduce [(row)]]]
		]

		all [
			empty? data
			either save [write file "" exit][return data]
		]

		if only [
			only: unset!
			either size > 1 [
				do compose/deep [
					remove-each [(row)] sort/skip/all data size [
						either only = reduce [(row)][true][only: reduce [(row)] false]
					]
				]
			][
				remove-each val sort data [
					either only = val [true][only: val false]
				]
			]
		]

		if merge [
			rowids: munge/part inner-block inner-size append munge/part keys 2 2 'rowid
			either 2 = length? keys [key: pick row first keys][
				key: make block! divide length? keys 2
				foreach col munge/part keys 2 1 [
					append key pick row col
				]
			]

			part: copy []
			foreach col cols [
				append part either col <= size [pick row col][
					compose [pick inner-block (col - size)]
				]
			]
			size: length? cols

			blk: copy []
			do compose/deep [
				foreach [(row)] data [
					all [
						rowid: select/skip rowids (either word? key [key][compose/deep [reduce [(key)]]]) (1 + divide length? keys 2)
						inner-block: skip head inner-block rowid - 1 * (inner-size)
						append blk reduce [(part)]
					]
				]
			]

			inner-block: head inner-block

			data: blk
		]

		if group [
			unless operation: any [
				all [find form having "count" 'count]
				all [find form having "flip" 'flip]
				all [find form having "max" 'max]
				all [find form having "min" 'min]
				all [find form having "sum" 'sum]
			][cause-error 'user 'message "Invalid group operation"]
			switch/default operation [
				count [
					i: 0
					blk: copy []
					either size = 1 [sort data][sort/skip/all data size]
					group: copy/part data size
					loop divide length? data size compose/deep [
						either group = copy/part data (size) [++ i][
							insert insert tail blk group i
							group: copy/part data (size)
							i: 1
						]
						data: skip data (size)
					]
					insert insert tail blk group i
					data: blk
					++ size
					append row operation
				]
				flip [
					blk: make block! length? data
					repeat col size compose/deep [
						i: col - (size)
						loop (divide length? data size) [
							append blk pick data i: i + (size)
						]
					]
					insert clear data blk
					size: divide length? data size
				]
			][
				i: 0
				either size = 1 [
					data: switch operation [
						max [copy/part maximum-of data 1]
						min [copy/part minimum-of data 1]
						sum [foreach val data [i: i + val] data: reduce [i]]
					]
					1
				][
					val: copy []
					blk: copy []
					sort/skip/all data size
					group: copy/part data size - 1
					loop divide length? data size compose/deep [
						either group = copy/part data (size - 1) [
							(either operation = 'sum [[i: i + pick data size]][[append val pick data size]])
						][
							insert insert tail blk group (
								switch operation [
									max [[first maximum-of val]]
									min [[first minimum-of val]]
									sum [[i]]
								]
							)
							group: copy/part data (size - 1)
							(either operation = 'sum [[i: pick data size]][[val: to block! pick data size]])
						]
						data: skip data (size)
					]
					insert insert tail blk group switch operation [
						max [first maximum-of val]
						min [first minimum-of val]
						sum [i]
					]
					data: blk
					poke row size operation
				]
			]
			blk: val: none
			all [block? having return munge/where data row having]
		]

		all [order sort/skip/all skip data either headings [size][0] size]

		also case [
			save [
				any [file file: %.]
				either find [%.xlsx %.xml] suffix? file [
					either headings [worksheet/new/save data size file][worksheet/new/no-header/save data size file]
				][
					blk: copy ""
					i: 1
					loop divide length? data size compose/deep [
						loop (size) [
							insert insert tail blk (
								either %.csv = suffix? file [
									[either find form val: pick data ++ i "," [ajoin [{"} val {"}]][val] ","]
								][
									[pick data ++ i "^-"]
								]
							)
						]
						poke blk length? blk #"^/"
					]
					either file = %. [parse/all blk "^/"][write file blk true]
				]
			]
			list [new-line/all/skip data true size]
			true [data]
		][data: blk: none recycle]
	]

	set 'read-pdf funct [	;	requires pdftotext.exe from http://www.foolabs.com/xpdf/download.html
		"Reads from a PDF file."
		file [file!]
		/lines "Handles data as lines"
	][
		any [exists? file cause-error 'access 'cannot-open file]
		call/wait ajoin [{"} base-path {pdftotext.exe" -nopgbrk -table "} to-local-file clean-path file {" $tmp.txt}]
		also either lines [
			remove-each line read/lines %$tmp.txt [empty? trim/tail line]
		][trim/tail read/string %$tmp.txt] delete %$tmp.txt
	]

	set 'read-xls funct [
		"Reads from an XLS file."
		file [file!]
		/sheet "Excel worksheet (default is 1)"
			number [integer! string! none!]
		/all "All worksheets"
		/lines "Handles data as lines"
	][
		either all [
			call-excel-vbs file 1 ajoin [{For I=1 to X.Worksheets.Count:X.ActiveWorkbook.Worksheets(I).SaveAs "} to-local-file file: join first split-path clean-path file %$tmp {"&I&".csv",6:Next:}]
			write/string file: join file %.csv ""
			repeat i to-integer first excel-metrics [
				write/string/append file read/string sheet: head insert skip tail copy file -4 i
				delete sheet
			]
		][
			call-excel-vbs file number ajoin [{X.ActiveWorkbook.Worksheets(} sheet? number {).SaveAs "} to-local-file file: join first split-path clean-path file %$tmp.csv {",6:}]
		]
		either exists? file [
			also either lines [
				remove-each line read/lines file ["," = unique trim line]
			][trim/tail read/string file] delete file
		][cause-error 'access 'cannot-open "$tmp.csv"]
	]

	set 'sqlcmd funct [
		"Execute a SQL Server statement."
		server [string!]
		database [string!]
		statement [string!]
		/key "Columns to convert to integer"
			columns [integer! block!]
		/headings "Keep column headings"
	][
		database: ajoin [server "/" database]
		any [columns columns: make block! 0]
		either headings [
			execute/key/headings database statement columns
		][
			execute/key database statement columns
		]
	]

	set 'unzip funct [
		"Uncompress a file into a folder of the same name (BETA)."
		file [file!]
		/only "Use current folder"
		/flatten "Move sub-folder(s) to current"
	][
		any [exists? file cause-error 'access 'cannot-open file]
		either only [
			path: first split-path clean-path file
		][
			make-dir path: replace clean-path file %.zip "/"
		]
		write %$tmp.vbs ajoin [{
			set F=CreateObject("Scripting.FileSystemObject")
			set S=CreateObject("Shell.Application")
			S.NameSpace("} to-local-file path {").CopyHere(S.NameSpace("} to-local-file clean-path file {").items)
			Set F=Nothing
			Set S=Nothing
		}]
		folders: remove-each folder read path [#"/" <> last folder]
		call/wait/shell "$tmp.vbs"
		delete %$tmp.vbs
		if flatten [
			foreach folder remove-each folder read path [#"/" <> last folder] [
				unless find folders folder [
					foreach file read join path folder [
						rename rejoin [path folder file] join %../ file
					]
					delete-dir join path folder
				]
			]
		]
		path
	]

	set 'worksheet funct [
		"Add a worksheet to the current workbook."
		data [block!]
		spec [integer! block!]
		/new "Start a new workbook"
		/no-header
		/sheet
			name [string!]
		/save "Save current workbook as an XML or XLSX file"
			file [file!]
	][
		either integer? spec [
			cols: spec
			insert/dup spec: make block! cols 100 cols
		][
			cols: length? spec
		]

		workbook: ""
		sheets: []

		all [
			any [new empty? workbook]
			insert clear workbook {<?xml version="1.0"?><?mso-application progid="Excel.Sheet"?><Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet" xmlns:html="http://www.w3.org/TR/REC-html40"><Styles><Style ss:ID="s1"><Interior ss:Color="#DDDDDD" ss:Pattern="Solid"/></Style><Style ss:ID="s2"><NumberFormat ss:Format="Long Date"/></Style><Style ss:ID="s3"><NumberFormat ss:Format="0%"/></Style></Styles>}
			clear sheets
			exit
		]

		any [integer? rows: divide length? data cols cause-error 'user 'message "column / data mismatch"]

		either sheet [trim/with to-xml-string name "/\"][name: reform ["Sheet" 1 + length? sheets]]
		either find sheets name [cause-error 'user 'message "Duplicate sheet name"][append sheets name]
		append workbook ajoin [{<Worksheet ss:Name="} name {"><Table>}]

		unless zero? rows [
		
			foreach width spec [
				append workbook ajoin [{<Column ss:Width="} width {"/>}]
			]

			no-header: either no-header [0][
				append workbook <Row>
				repeat i cols [
					append workbook ajoin [<Cell ss:StyleID="s1"><Data ss:Type="String"> first data </Data></Cell>]
					data: next data
				]
				append workbook </Row>
				1
			]

			loop rows - no-header [
				append workbook <Row>
				loop cols [
					val: first data
					append workbook case [
						all [string? val parse/all val [some digit opt [["." | ","] some digit] "%" end] ] [
							ajoin [<Cell ss:StyleID="s3"><Data ss:Type="Number"> to decimal! head remove back tail val </Data></Cell>]
						]	;percent support (added by endo)
						number? val [
							ajoin [<Cell><Data ss:Type="Number"> val </Data></Cell>]
						]
						date? val [
							ajoin [
								<Cell ss:StyleID="s2"><Data ss:Type="DateTime">
								head insert skip head insert skip form (val/year * 10000) + (val/month * 100) + val/day 6 "-" 4 "-"
								</Data></Cell>
							]
						]
						all [string? val "=" = copy/part val 1][
							ajoin [{<Cell ss:Formula="} val {"><Data ss:Type="Number"></Data></Cell>}]
						]
						true [
							either any [none? val empty? val: trim form val][<Cell />][
								ajoin [<Cell><Data ss:Type="String"> to-xml-string val </Data></Cell>]
							]
						]
					]
					data: next data
				]
				append workbook </Row>
			]

			data: head data
		]

		append workbook "</Table></Worksheet>"

		if save [
			append workbook </Workbook>
			switch suffix? file [
				%.xml [
					foreach tag ["<Workbook" "<?mso-application" "<Style" </Styles> "<Worksheet" <Table> "<Column" <Row> "<Cell" </Row> </Table> </Worksheet> </Workbook>][
						replace/all workbook tag either tag = "<Cell" [join "^/^-" tag][join "^/" tag]
					]
					write file turkish-to-utf workbook
				]
				%.xlsx [
					write %$tmp.xml turkish-to-utf workbook
					also call-excel-vbs %$tmp.xml 1 ajoin [{W.SaveAs "} to-local-file clean-path file {",51:}] delete %$tmp.xml
					any [exists? file cause-error 'access 'cannot-open file]
				]
			]
			clear workbook
			clear sheets
		]

		exit
	]
]
