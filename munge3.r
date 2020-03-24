Rebol [
	Title:		"Munge functions"
	Owner:		"Ashley G Truter"
	Version:	3.0.1
	Date:		18-Apr-2017
	Purpose:	"Extract and manipulate tabular values in blocks, delimited files and database tables."
	Licence:	"MIT. Free for both commercial and non-commercial use."
	Tested: {
		3.0.99.3.3			R3/64-bit	atronixengineering.com
		0.6.2				RED/32-bit	www.red-lang.org
	}
	Usage: {
		append-column		Append a column of values to a block.
		ascii-file?			Returns TRUE if file is ASCII.
		average-of			Average of values in a block.
		call-oledb			Call OLEDB via PowerShell returning STDOUT.
		call-out			Call OS command returning STDOUT.
		check				Verify data structure.
		cols?				Number of columns in a delimited file or string.
		delimiter?			Probable delimiter, with priority given to tab, bar, tilde, semi-colon then comma.
		digit				digit is a bitset! of value: make bitset! #{000000000000FFC0}
		digits?				Returns TRUE if data not empty and only contains digits.
		distinct			Remove duplicate and empty rows.
		enblock				Convert a block of values to a block of row blocks.
		export				Export words to global context.
		fields?				Column names in a delimited file or string.
		first-line			Returns the first non-empty line of a file.
		flatten				Flatten nested block(s).
		latin1-to-utf8		Latin1 binary to UTF-8 string conversion.
		letter				letter is a bitset! of value: make bitset! #{00000000000000007FFFFFE07FFFFFE0}
		letters?			Returns TRUE if data only contains letters.
		like				Finds a value in a series, expanding * (any characters) and ? (any one character), and returns TRUE if found.
		list				Sets the new-line marker to end of block.
		load-dsv			Parses delimiter-separated values into row blocks.
		load-excel			Loads an Excel file.
		max-of				Returns the largest value in a series.
		merge				Join outer block to inner block on primary key.
		min-of				Returns the smallest value in a series.
		mixedcase			Converts string of characters to mixedcase.
		munge				Load and/or manipulate a block of tabular (column and row) values.
		oledb-file?			Returns true for an Excel or Access file.
		read-pdf			Reads from a PDF file.
		read-string			Read string(s) from a text file.
		remove-column		Remove a column of values from a block.
		replace-deep		Replaces all occurences of a search value with the new value in a block or nested block.
		rows?				Number of rows in a delimited file or string.
		sheets?				Excel sheet names.
		split-line			Splits and returns line of text as a block.
		sqlcmd				Execute a SQL Server statement.
		sqlite				Execute a SQLite statement.
		sum-of				Sum of values in a block.
		to-column-alpha		Convert numeric column reference to an alpha column reference.
		to-column-number	Convert alpha column reference to a numeric column reference.
		to-hash				Convert block! to map!.
		to-rebol-date		Convert a string date to a Rebol date.
		to-rebol-time		Convert a string date/time to a Rebol time.
		to-string-date		Convert a string or Rebol date to a YYYY-MM-DD string.
		to-string-time		Convert a string or Rebol time to a HH.MM.SS string.
		write-dsv			Write block(s) of values to a delimited text file.
		write-excel			Write block(s) of values to an Excel file.
	}
]

Red []

either rebol [
	foreach word [put][
		all [value? word print [word "already defined!"]]
	]

	all [block? system/platform system/platform: system/platform/1]

	put: function [
		"Replaces the value following a key, and returns the map"
		map [map!]
		key
		value
	][
		append map reduce [key value]
	]
][
	foreach word [date! delete delete-dir deline invalid-utf? reform to-date][
		all [value? word print [word "already defined!"]]
	]

	date!: :string!

	delete: function [
		"Deletes a file"
		file [file!]
	] compose/deep [
		either exists? file [
			call/wait rejoin [(either system/platform = 'Windows [{del "}][{rm "}]) to-local-file file {"}]
		][
			cause-error 'access 'cannot-open reduce [file]
		]
	]

	delete-dir: function [
		"Deletes a directory including all files and subdirectories"
		file [file!]
	] compose/deep [
		either exists? file [
			call/wait rejoin [(either system/platform = 'Windows [{rd /s /q "}][{rm -r "}]) to-local-file file {"}]
		][
			cause-error 'access 'cannot-open reduce [file]
		]
	]

	deline: function [
		string [any-string!]
		/lines "Return block of lines (works for LF, CR, CR-LF endings) (no modify)"
		"Converts string terminators to standard format, e.g. CRLF to LF"
	][
		replace/all string crlf lf
		either lines [split string lf][string]
	]

	invalid-utf?: function [
		"Checks UTF encoding; if correct, returns none else position of error"
		binary [binary!]
	] compose [
		find binary (make bitset! [192 193 245 - 255])
	]

	unless value? 'now* [
		now*: :now
		now: function [
			"Returns date and time"
			/year "Returns year only"
			/month "Returns month only"
			/day "Returns day of the month only"
			/time "Returns time only"
			/date "Returns date only"
			/weekday "Returns day of the week as integer (Monday is day 1)"
			/precise "High precision time"
		] compose [
			t: either precise [now*/time/precise][now*/time]
			call/wait/output (either system/platform = 'Windows ["date /t"]["date +%Y-%m-%d"]) s: make string! 16
			d: to-date any [find s " " s]
			case [
				year	[d/year]
				month	[d/month]
				day		[d/day]
				date	[d/date]
				weekday	[copy/part s 3]
				time	[t]
				true	[rejoin [d/date "/" t]]
			]
		]
	]

	reform: function [
		"Forms a reduced block and returns a string"
		value "Value to reduce and form"
	][
		form reduce value
	]

	to-date: function [
		"Converts to date! value"
		date [string!]
	][
		d: ctx-munge/split-line date either find date "-" [#"-"][#"/"]
		all [
			4 = length? d/1
			d: reduce [d/3 d/2 d/1]
		]
		d/1: to integer! d/1
		d/2: to integer! d/2
		d/3: to integer! d/3
		all [
			100 > d/3
			d/3: d/3 + either d/3 < 68 [2000][1900]
		]
		make object! [
			date:	rejoin [d/1 "-" copy/part pick system/locale/months d/2 3 "-" d/3]
			year:	d/3
			month:	d/2
			day:	d/1
		]
	]
]

ctx-munge: context [

	append-column: function [
		"Append a column of values to a block"
		block [block!]
		value
		/header heading
	][
		unless empty? block [
			foreach row block [
				append row value
			]
			all [header poke block/1 length? block/1 heading]
		]
		block
	]

	ascii-file?: function [
		"Returns TRUE if file is ASCII"
		file [file!]
	] either rebol [[
		ascii? read/string/part file 1024
	]][compose [
		not find read/binary/part file 1024 (charset [128 - 255])
	]]

	average-of: function [
		"Average of values in a block"
		block [block!]
	][
		all [empty? block return none]
		total: 0
		foreach value block [total: total + value]
		total / length? block
	]

	call-oledb: function [
		"Call OLEDB via PowerShell returning STDOUT"
		;	Excel, using https://www.microsoft.com/en-au/download/details.aspx?id=13255
		file [file!]
		cmd [string!]
		/hdr "First row contains columnnames not data"
	][
		any [exists? file cause-error 'access 'cannot-open reduce [file]]
		all [find file %' cause-error 'user 'message ["file name contains an invalid ' character"]]
		trim/tail trim call-out rejoin [
			either rebol ["powershell "]["C:\Windows\SysNative\WindowsPowerShell\v1.0\powershell.exe "]
			{-nologo -noprofile -command "}
			{$o=New-Object System.Data.OleDb.OleDbConnection('Provider=Microsoft.ACE.OLEDB.12.0;Data Source=''} to-local-file clean-path file {''}
			either %.accdb = suffix? file ["');"][rejoin [{;Extended Properties=''Excel 12.0 Xml;HDR=} either hdr ["YES"]["NO"] {;IMEX=1;Mode=Read''');}]]
			cmd
			either #";" = last cmd [
				rejoin [
					{$s.Connection=$o;}
					{$t=New-Object System.Data.DataTable;}
					{$t.Load($s.ExecuteReader());}
					{$o.Close();}
					{$t|ConvertTo-CSV -Delimiter `t -NoTypeInformation}
				]
			][""]
			{"}
		]
	]

	call-out: function [
		"Call OS command returning STDOUT"
		cmd [string!]
	][
		call/wait/output/error cmd stdout: make binary! 65536 stderr: make string! 1024
		any [empty? stderr cause-error 'user 'message reduce [trim/lines stderr]]
		read-string stdout
	]

	check: function [
		"Verify data structure"
		data [block!]
	][
		cols: length? data/1
		i: 1
		foreach row data [
			any [block? row cause-error 'user 'message reduce [reform ["Row" i "expected block but found" type? row]]]
			all [zero? length? row cause-error 'user 'message reduce [reform ["Row" i "empty"]]]
			any [cols = length? row cause-error 'user 'message reduce [reform ["Row" i "expected" cols "column(s) but found" length? row]]]
			all [block? row/1 row cause-error 'user 'message reduce [reform ["Row" i "did not expect first column to be a block"]]]
			i: i + 1
		]
		true
	]

	cols?: function [
		"Number of columns in a delimited file or string"
		data [file! binary! string!]
		/with
			delimiter [char!]
		/sheet {Excel worksheet name (default is "Sheet 1")}
			name [string! word!]
	][
		length? either sheet [fields?/sheet data name] [
			either delimiter [
				fields?/with data delimiter
			][
				fields? data
			]
		]
	]

	delimiter?: function [
		"Probable delimiter, with priority given to tab, bar, tilde, semi-colon then comma"
		data [file! string!]
	][
		data: first-line data
		case [
			find data tab [tab]
			find data #"|" [#"|"]
			find data #"~" [#"~"]
			find data #";" [#";"]
			true [#","]
		]
	]

	digit: charset [#"0" - #"9"]

	digits?: function [
		"Returns TRUE if data not empty and only contains digits"
		data [string! binary!]
	] compose/deep [
		all [not empty? data not find data (complement digit)]
	]

	distinct: function [
		"Remove duplicate and empty rows"
		data [block!]
	][
		old-row: none
		remove-each row sort data [
			any [
				all [series? row/1 empty? row/1 1 = length? unique row]
				all [none? row/1 1 = length? unique row]
				either row == old-row [true][old-row: row false]
			]
		]
		data
	]

	enblock: function [
		"Convert a block of values to a block of row blocks"
		data [block!]
		cols [integer!]
	][
		all [block? data/1 return data]
		any [integer? rows: divide length? data cols cause-error 'user 'message ["Cols not a multiple of length"]]
		i: 1
		loop rows [
			change/part/only at data i copy/part at data i cols cols
			i: i + 1
		]
		data
	]

	export: function [
		"Export words to global context"
		words [block!] "Words to export"
	][
		foreach word words [
			do compose [(to-set-word word) (to-get-word in self word)]
		]
		words
	]

	fields?: function [
		"Column names in a delimited file or string"
		data [file! string!]
		/with
			delimiter [char!]
		/sheet {Excel worksheet name (default is "Sheet 1")}
			name [string! word!]
	] compose/deep [
		either any [string? data not oledb-file? data] [
			either data: first-line data [
				split-line data any [delimiter delimiter? data]
			][
				make block! 0
			]
		][
			sheet: any [name first sheets? data]
			unless %.accdb = suffix? data [
				sheet: rejoin [sheet "$"]
				all [find sheet " " insert sheet "''" append sheet "''"]
			]
			remove/part deline/lines call-oledb/hdr data rejoin [
				{$o.Open();$o.GetSchema('Columns')|where TABLE_NAME -eq '} sheet {'|sort ORDINAL_POSITION|select COLUMN_NAME;$o.Close()}
			] 2
		]
	]

	first-line: function [
		"Returns the first non-empty line of a file"
		data [file! string!]
	] compose/deep [
		foreach line deline/lines either file? data [
			latin1-to-utf8 (either rebol [[read/part]][[read/binary/part]]) data 4096
		][
			copy/part data 4096
		][
			any [
				empty? line
				return line
			]
		]
		copy ""
	]

	flatten: function [ ; http://www.rebol.org/view-script.r?script=flatten.r
		"Flatten nested block(s)"
		data [block!]
	][
		all [empty? data return data]
		result: copy []
		foreach row data [
			append result row
		]
	]

	latin1-to-utf8: function [ ; http://stackoverflow.com/questions/21716201/perform-file-encoding-conversion-with-rebol-3
		"Latin1 binary to UTF-8 string conversion"
		data [binary!]
	][
		mark: data
		while [mark: find mark #{C2A0}][	; replace char 160 with space - http://www.adamkoch.com/2009/07/25/white-space-and-character-160/
			change/part mark #{20} 2
		]
		mark: data
		while [mark: invalid-utf? mark][	; replace latin1 with utf
			change/part mark to char! mark/1 1
		]
		trim/with to string! data null		; remove #"^@"
	]

	letter: charset [#"A" - #"Z" #"a" - #"z"]

	letters?: function [
		"Returns TRUE if data only contains letters"
		data [string! binary!]
	] compose [
		not find data (complement letter)
	]

	like: function [ ; http://stackoverflow.com/questions/31612164/does-anyone-have-an-efficient-r3-function-that-mimics-the-behaviour-of-find-any
		"Finds a value in a series, expanding * (any characters) and ? (any one character), and returns TRUE if found"
		series [any-string!] "Series to search"
		value [any-string! block!] "Value to find"
	] compose [
		all [empty? series return none]
		literal: (complement charset "*?")
		value: collect [
			parse value [
				end (keep [return (none)]) |
				some #"*" end (keep [to end]) |
				some [
					#"?" (keep 'skip) |
					copy part some literal (keep part) |
					some #"*" any [#"?" (keep 'skip)] opt [copy part some literal (keep 'thru keep part)]
				]
			]
		]
		parse series [some [result: value (return true)]]
	]

	list: function [
		"Sets the new-line marker to end of block"
		data [block!]
	][
		new-line/all data true
	]

	load-dsv: function [ ; http://www.rebol.org/view-script.r?script=csv-tools.r
		"Parses delimiter-separated values into row blocks"
		source [file! binary! string!]
		/part "Offset position(s) to retrieve"
			columns [block! integer!]
		/preserve "keep line breaks and extra spaces"
		/ignore "Ignore truncated row errors"
		/where "Expression that can reference columns as f1, f2, etc"
			condition [block!]
		/with "Alternate delimiter (default is tab, bar then comma)"
			delimiter [char!]
	][
		source: case [
			file? source	[either ascii-file? source [read-string source][cause-error 'user 'message ["Cannot use load-dsv with binary file"]]]
			binary? source	[latin1-to-utf8 source]
			string? source	[deline source]
		]

		any [delimiter delimiter: delimiter? source]

		value: either preserve [[
			{"} (clear v) x: to [{"} | end] y: (append/part v x y)
			any [{"} x: {"} to [{"} | end] y: (append/part v x y)]
			[{"} x: to [delimiter | lf | end] y: (append/part v x y) | end]
			(append row copy v) |
			copy x to [delimiter | lf | end] (append row x)
		]][[
			{"} (clear v) x: to [{"} | end] y: (append/part v x y)
			any [{"} x: {"} to [{"} | end] y: (append/part v x y)]
			[{"} x: to [delimiter | lf | end] y: (append/part v x y) | end]
			(append row trim/lines copy v) |
			copy x to [delimiter | lf | end] (append row trim/lines x)
		]]

		cols: cols?/with source delimiter

		columns: either part [
			all [
				not ignore
				cols < i: max-of columns: to block! columns
				cause-error 'user 'message reduce [reform ["Expected a value in column" i "but only found" cols "columns"]]
			]
			part: copy/deep [reduce []]
			foreach col to block! columns [
				append part/2 compose [(append to path! 'row col)]
			]
			part
		]['row]

		v: make string! 32

		append-row: function [
			row [block!]
		] compose/deep either any [part where][[
			all [
				(either where [condition][])
				(either ignore [][compose/deep [any [(cols) = len: length? row cause-error 'user 'message reduce [reform ["Expected" (cols) "values but found" len "on line" line]]]]])
				append/only blk (columns)
			]
		]][[
			(either ignore [][compose/deep [all [(cols) <> len: length? row cause-error 'user 'message reduce [reform ["Expected" (cols) "values but found" len "on line" line]]]]])
			append/only blk row
		]]

		line: 0
		blk: copy []
		parse source [
			any [
				not end (row: make block! cols)
				value
				any [delimiter value] [lf | end] (line: line + 1 all [[""] <> unique row append-row row])
			]
		]

		blk
	]

	load-excel: function [
		"Loads an Excel file"
		file [file!]
		sheet [integer! word! string!]
		/part "Offset position(s) / columns(s) to retrieve"
			columns [block! integer!]
		/where "Expression that can reference columns as F1, F2, etc"
			condition [string!]
		/distinct
		/string
	][
		any [exists? file cause-error 'access 'cannot-open reduce [file]]
		all [find condition "'" condition: replace/all copy condition "'" "''"]
		either part [
			part: copy ""
			foreach i to block! columns [append part rejoin ["F" i ","]]
			remove back tail part
		] [part: "*"]

		unless any [integer? sheet %.accdb = suffix? file] [
			sheet: rejoin [sheet "$"]
			all [find sheet " " insert sheet "''" append sheet "''"]
		]

		stdout: call-oledb file rejoin [
			{$o.Open();$s=New-Object System.Data.OleDb.OleDbCommand('}
			"SELECT " either distinct ["DISTINCT "][""] part
			" FROM [" either integer? sheet [rejoin ["'+$o.GetSchema('Tables').rows[" sheet - 1 "].TABLE_NAME+'"]][sheet] "]"
			either where [reform [" WHERE" condition]][""]
			{');}
		]
		either string [stdout] [remove load-dsv/with stdout #"^-"]
	]

	max-of: function [
		"Returns the largest value in a series"
		series [series!] "Series to search"
	][
		val: series/1
		foreach v series [val: max val v]
		val
	]

	merge: function [
		"Join outer block to inner block on primary key" 
		outer [block!] "Outer block"
		key1 [integer!]
		inner [block!] "Inner block to index"
		key2 [integer!]
		columns [block!] "Offset position(s) to retrieve in merged block"
		/default "Use none on inner block misses"
	][
		;	build rowid map of inner block
		map: make map! length? inner
		i: 0
		foreach row inner [
			put map pick row key2 i: i + 1
		]
		;	build column picker
		code: copy []
		foreach column columns [
			append code compose [pick row (column)]
		]
		;	iterate through outer block
		blk: make block! length? outer
		do compose/deep [
			either default [
				foreach row outer [
					all [i: select map pick row key1 append row pick inner i]
					append/only blk reduce [(code)]
				]
			][
				foreach row outer [
					all [
						i: select map pick row key1
						append row pick inner i
						append/only blk reduce [(code)]
					]
				]
			]
		]

		blk
	]

	min-of: function [
		"Returns the smallest value in a series"
		series [series!] "Series to search"
	][
		val: series/1
		foreach v series [val: min val v]
		val
	]

	mixedcase: function [
		"Converts string of characters to mixedcase"
		string [string!]
	][
		uppercase/part lowercase string 1
		foreach char [#"'" #" " #"-" #"." #","][
			all [find string char string: next find string char mixedcase string]
		]
		string: head string
	]

	munge: function [
		"Load and/or manipulate a block of tabular (column and row) values"
		data [block!]
		/update "Offset/value pairs (returns original block)"
			action [block!]
		/delete "Delete matching rows (returns original block)"
		/part "Offset position(s) to retrieve"
			columns [block! integer!]
		/where "Expression that can reference columns as f1, f2, etc"
			condition
		/group "One of count, max, min or sum"
			having [word! block!] "Word or expression that can reference the initial result set column as count, max, etc"
	][
		all [empty? data return data]

		if all [where not block? condition] [ ; http://www.rebol.org/view-script.r?script=binary-search.r
			lo: 1
			hi: rows: length? data
			mid: to integer! hi + lo / 2
			while [hi >= lo][
				if condition = key: first pick data mid [
					lo: hi: mid
					while [all [lo > 1 condition = first pick data lo - 1]] [lo: lo - 1]
					while [all [hi < rows condition = first pick data hi + 1]] [hi: hi + 1]
					break
				]
				either condition > key [lo: mid + 1][hi: mid - 1]
				mid: to integer! hi + lo / 2
			]
			all [
				lo > hi
				return either any [update delete] [data] [make block! 0]
			]
			rows: hi - lo + 1
			case [
				update	[]
				delete	[return head remove/part at data lo rows]
				true	[data: copy/part at data lo rows where: none]
			]
		]

		case [
			update [
				blk: copy []
				foreach [col val] action [
					append blk compose [
						(append to set-path! 'row col) (either all [word? val #"!" = last form val] [compose [to (val) (append to path! 'row col)]] [val])
					]
				]
				either rebol [
					either block? condition [
						foreach row data compose/deep [
							all [
								(condition)
								(blk)
							]
						]
					][
						foreach row either where [copy/part at data lo rows][data] compose [
							(blk)
						]
					]
				][
					either block? condition [
						foreach row data bind compose/deep [
							all [
								(condition)
								(blk)
							]
						] 'row
					][
						foreach row either where [copy/part at data lo rows][data] bind compose [
							(blk)
						] 'row
					]
				]
				return data
			]
			delete [
				either where [
					either rebol [
						remove-each row data compose/only [all (condition)]
					][
						remove-each row data bind compose/only [all (condition)] 'row
					]
				][
					clear data
				]
				return data
			]
			any [part where] [
				columns: either part [
					part: copy/deep [reduce []]
					foreach col to block! columns [
						append part/2 compose [(append to path! 'row col)]
					]
					part
				]['row]
				blk: copy []
				foreach row data compose [
					(
						either where [
							either rebol [
								compose/deep [all [(condition) append/only blk (columns)]]
							][
								bind compose/deep [all [(condition) append/only blk (columns)]] 'row
							]
						][
							compose [append/only blk (columns)]
						]
					)
				]
				all [empty? blk return blk]
				data: blk
			]
		]

		if group [
			words: flatten to block! having
			operation: any [
				all [find words 'avg 'average-of]
				all [find words 'count 'count]
				all [find words 'max 'max-of]
				all [find words 'min 'min-of]
				all [find words 'sum 'sum-of]
				cause-error 'user 'message ["Invalid group operation"]
			]
			case [
				operation = 'count [
					i: 0
					blk: copy []
					group: copy first sort data
					foreach row data [
						either group = row [i: i + 1][
							append group i
							append/only blk group
							group: row
							i: 1
						]
					]
					append group i
					append/only blk group
				]
				1 = length? data/1 [
					return do compose [(operation) flatten data]
				]
				true [
					val: copy []
					blk: copy []
					group: copy/part first sort data len: -1 + length? data/1
					foreach row data compose/deep [
						either group = copy/part row (len) [append val last row][
							append group (operation) val
							append/only blk group
							group: copy/part row (len)
							append val: copy [] last row
						]
					]
					append group do compose [(operation) val]
					append/only blk group
				]
			]
			data: blk

			all [
				block? having
				replace-deep having operation append to path! 'row length? data/1
				return munge/where data having
			]
		]

		data
	]

	oledb-file?: function [
		"Returns true for an Excel or Access file"
		file [file!]
	][
		all [
			suffix? file
			find [%.acc %.xls] copy/part suffix? file 4
			exists? file
			not ascii-file? file
		]
	]

	read-pdf: function [ ; requires pdftotext.exe from http://www.foolabs.com/xpdf/download.html
		"Reads from a PDF file"
		file [file!]
		/lines "Handles data as lines"
	] compose/deep [
		any [exists? file cause-error 'access 'cannot-open reduce [file]]
		call/wait rejoin [{"} (to-local-file system/options/path) {pdftotext" -nopgbrk -table "} to-local-file clean-path file {" tmp.txt}]
		also either lines [read-string/lines %tmp.txt] [trim/tail read-string %tmp.txt] delete %tmp.txt
	]

	read-string: function [
		"Read string(s) from a text file"
		data [file! binary!]
		/lines "Convert to block of strings"
	] compose/deep [
		either file? data [
			either lines [
				deline/lines latin1-to-utf8 (either rebol [[read data]][[read/binary data]])
			][
				deline latin1-to-utf8 (either rebol [[read data]][[read/binary data]])
			]
		][
			either lines [deline/lines latin1-to-utf8 data][deline latin1-to-utf8 data]
		]
	]

	remove-column: function [
		"Remove a column of values from a block"
		block [block!]
		index [integer!]
	][
		foreach row block [
			remove at row index
		]
		block
	]

	replace-deep: function [
		"Replaces all occurences of a search value with the new value in a block or nested block"
		data [block!] "Block to replace within (modified)"
		search "Value to be replaced"
		new 
	][
		replace/all data search new
		foreach value data [
			all [block? value replace-deep value search new]
		]
		data
	]

	rows?: function [
		"Number of rows in a delimited file or string"
		data [file! binary! string!]
	][
		either any [
			all [file? data zero? size? data]
			empty? data
		][0][
			i: 1
			parse either file? data [read data][data] [
				any [thru newline (i: i + 1)]
			]
			i
		]
	]

	sheets?: function [
		"Excel sheet names"
		file [file!]
	][
		blk: remove/part deline/lines call-oledb file {$o.Open();$o.GetSchema('Tables')|where TABLE_TYPE -eq "TABLE"|SELECT TABLE_NAME;$o.Close()} 2
		unless %.accdb = suffix? file [
			foreach s blk [trim/with s "'"]
			remove-each s blk [#"$" <> last s]
			foreach s blk [remove back tail s]
		]
		blk
	]

	split-line: function [
		"Splits and returns line of text as a block"
		line [string!]
		delimiter [char!]
	] either rebol [[
		foreach s row: split line delimiter [trim/lines trim/with s {"}]
		row
	]][[
		foreach s row: split line delimiter [trim/lines trim/with s {"}]
		all [delimiter = last line append row copy ""]
		all [[""] = row clear row]
		row
	]]

	sqlcmd: function [
		"Execute a SQL Server statement"
		server [string!]
		database [string!]
		statement [string!]
		/key "Columns to convert to integer"
			columns [integer! block!]
		/headings "Keep column headings"
		/string
	][
		stdout: call-out reform compose ["sqlcmd -X -S" server "-d" database "-I -Q" rejoin [{"} statement {"}] {-W -w 65535 -s"^-"} (either headings [][{-h -1}])]
		all [string return stdout]
		all [any [empty? stdout #"^/" = first stdout] return make block! 0]
		all [like first-line stdout "Msg*,*Level*,*State*,*Server" cause-error 'user 'message reduce [trim/lines find stdout "Line"]]

		either "^/(" = copy/part stdout 2 [
			trim/with stdout "()^/"
		][
			stdout: copy/part stdout find stdout "^/^/("
			all [find ["" "NULL"] stdout return make block! 0]

			stdout: load-dsv/with stdout #"^-"

			all [headings remove skip stdout 1]

			foreach row stdout [
				foreach val row [
					all ["NULL" == trim val clear val]
				]
			]

			remove-each row stdout [
				all [series? row/1 empty? row/1 1 = length? unique row]
			]

			all [
				key
				foreach row skip stdout either headings [1][0] [
					foreach i to block! columns [
						poke row i to integer! pick row i
					]
				]
			]

			stdout
		]
	]

	sqlite: function [
		"Execute a SQLite statement"
		database [file!]
		statement [string!]
		/key "Columns to convert to integer"
			columns [integer! block!]
		/headings "Keep column headings"
		/string
	][
		stdout: call-out reform compose [{sqlite3 -separator "^-"} (either headings ["-header"][]) to-local-file database rejoin [{"} trim/lines statement {"}]]
		all [string return stdout]
		all [find ["" "^/"] stdout return make block! 0]
		stdout: load-dsv/with stdout #"^-"

		all [
			key
			foreach row skip stdout either headings [1][0] [
				foreach i to block! columns [
					poke row i to integer! pick row i
				]
			]
		]

		stdout
	]

	sum-of: function [
		"Sum of values in a block"
		block [block!]
	][
		all [empty? block return none]
		total: 0
		foreach value block [total: total + value]
	]

	to-column-alpha: function [
		"Convert numeric column reference to an alpha column reference"
		number [integer!] "Column number between 1 and 702"
	][
		any [positive? number cause-error 'user 'message ["Positive number expected"]]
		any [number <= 702 cause-error 'user 'message ["Number cannot exceed 702"]]
		either number <= 26 [form #"@" + number][
			rejoin [
				#"@" + to integer! number - 1 / 26
				either zero? r: mod number 26 ["Z"][#"@" + r]
			] 
		]
	]

	to-column-number: function [
		"Convert alpha column reference to a numeric column reference"
		alpha [word! string! char!]
	][
		any [find [1 2] length? alpha: uppercase form alpha cause-error 'user 'message ["One or two letters expected"]]
		any [find letter last alpha cause-error 'user 'message ["Valid characters are A-Z"]]
		minor: subtract to integer! last alpha: uppercase form alpha 64
		either 1 = length? alpha [minor] [
			any [find letter alpha/1 cause-error 'user 'message ["Valid characters are A-Z"]]
			(26 * subtract to integer! alpha/1 64) + minor
		]
	]

	to-hash: function [
		"Convert block! to map!"
		data [block!]
	][
		map: make map! length? data
		foreach value flatten data [
			put map value 0
		]
		map
	]

	to-rebol-date: function [
		"Convert a string date to a Rebol date"
		date [string!]
		/mdy "Month/Day/Year format"
		/ydm "Year/Day/Month format"
		/day "Day precedes date"
	] compose [
		all [day date: copy date remove/part date next find date " "]
		date: (either rebol [[parse date "/- "]][[split date charset "/- "]])
		to-date case [
			mdy		[rejoin [date/2 "-" date/1 "-" date/3]]
			ydm		[rejoin [date/2 "-" date/3 "-" date/1]]
			true	[rejoin [date/1 "-" date/2 "-" date/3]]
		]
	]

	to-rebol-time: function [
		"Convert a string date/time to a Rebol time"
		time [string!]
	] either rebol [[
		to time! trim/all copy back back find time ":"
	]][[
		either find time "PM" [
			time: to time! time
			all [time/1 < 13 time/1: time/1 + 12]
			time
		][
			to time! time
		]
	]]

	to-string-date: function [
		"Convert a string or Rebol date to a YYYY-MM-DD string"
		date [string! date!]
		/mdy "Month/Day/Year format"
		/ydm "Year/Day/Month format"
	] compose/deep [
		all [
			string? date
			date: to-date case [
				mdy		[date: (either rebol [[parse date "/- "]][[split date charset "/- "]]) rejoin [date/2 "/" date/1 "/" date/3]]
				ydm		[date: (either rebol [[parse date "/- "]][[split date charset "/- "]]) rejoin [date/2 "/" date/3 "/" date/1]]
				true	[date]
			]
		]
		rejoin [date/year "-" next form 100 + date/month "-" next form 100 + date/day]
	]

	to-string-time: function [
		"Convert a string or Rebol time to a HH.MM.SS string"
		time [string! date! time!]
	] either rebol [[
		all [
			string? time
			time: to time! trim/all copy time
		]
		rejoin [next form 100 + time/hour "." next form 100 + time/minute "." next form 100 + time/second]
	]][[
		if string? time [
			PM?: find time "PM"
			time: to time! time
			all [PM? time/1 < 13 time/1: time/1 + 12]
		]
		rejoin [next form 100 + time/hour "." next form 100 + time/minute "." next form 100 + to integer! time/second]
	]]

	write-dsv: function [
		"Write block(s) of values to a delimited text file"
		file [file!] "csv or tab-delimited text file"
		data [block!]
	][
		s: copy ""
		foreach row data compose/deep [
			foreach value row [
				append s (
					either %.csv = suffix? file [
						[rejoin [either any [find val: trim/with form value {"} "," find val lf] [rejoin [{"} val {"}]][val] ","]]
					][
						[rejoin [value "^-"]]
					]
				)
			]
			any [empty? s poke s length? s newline]
		]
		write file s
	]

	write-excel: function [ ; http://officeopenxml.com/anatomyofOOXML-xlsx.php
		"Write block(s) of values to an Excel file"
		file [file!]
		data [block!] "Name [string!] Data [block!] Widths [block!] records"
		/filter "Add auto filter"
	] compose [
		any [%.xlsx = suffix? file cause-error 'user 'message ["not a valid .xlsx file extension"]]
		all [exists? path: %tmp/ cause-error 'user 'message ["tmp/ already exists"]]

		make-dir/deep rejoin [path %_rels]
		make-dir/deep rejoin [path %xl/_rels]
		make-dir/deep rejoin [path %xl/worksheets]

		xml-content-types: copy ""
		xml-workbook: copy ""
		xml-workbook-rels: copy ""
		xml-version: {<?xml version="1.0" encoding="UTF-8" standalone="yes"?>}
		sheet-number: 1

		foreach [sheet-name block spec] data [
			unless empty? block [
				width: length? spec

				append xml-content-types rejoin [{<Override PartName="/xl/worksheets/sheet} sheet-number {.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>}]
				append xml-workbook rejoin [{<sheet name="} sheet-name {" sheetId="} sheet-number {" r:id="rId} sheet-number {"/>}]
				append xml-workbook-rels rejoin [{<Relationship Id="rId} sheet-number {" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet} sheet-number {.xml"/>}]

				;	%xl/worksheets/sheet<n>.xml

				blk: rejoin [
					xml-version
					{<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
						<cols>}
				]
				repeat i width [
					append blk rejoin [{<col min="} i {" max="} i {" width="} pick spec i {"/>}]
				]
				append blk "</cols><sheetData>"
				foreach row block [
					append blk "<row>"
					foreach value row [
						append blk case [
							number? value [
								rejoin ["<c><v>" value "</v></c>"]
							]
							#"=" = first value: form value [
								rejoin ["<c><f>" next value "</f></c>"]
							]
							true [
								foreach [char code][
									"&"		"&amp;"
									"<"		"&lt;"
									">"		"&gt;"
									{"}		"&quot;"
									{'}		"&apos;"
									"^/"	"&#10;"
								][replace/all value char code]
								rejoin [{<c t="inlineStr"><is><t>} value "</t></is></c>"]
							]
						]
					]
					append blk "</row>"
				]
				append blk {</sheetData>}
				all [filter append blk rejoin [{<autoFilter ref="A1:} to-column-alpha width length? block {"/>}]]
				append blk {</worksheet>}
				write rejoin [path %xl/worksheets/sheet sheet-number %.xml] blk

				sheet-number: sheet-number + 1
			]
		]

		write rejoin [path %"[Content_Types].xml"] rejoin [
			xml-version
			{<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
				<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
				<Default Extension="xml" ContentType="application/xml"/>
				<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>}
				xml-content-types
			{</Types>}
		]

		write rejoin [path %_rels/.rels] rejoin [
			xml-version
			{<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
				<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
			</Relationships>}
		]

		write rejoin [path %xl/workbook.xml] rejoin [
			xml-version
			{<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x15" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">
				<workbookPr defaultThemeVersion="153222"/>
				<sheets>}
					xml-workbook
				{</sheets>
			</workbook>}
		]

		write rejoin [path %xl/_rels/workbook.xml.rels] rejoin [
			xml-version
			{<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">}
				xml-workbook-rels
			{</Relationships>}
		]

		;	create xlsx file

		all [exists? file delete file]

		call/wait (either system/platform = 'Windows [[
			rejoin [ ; http://stackoverflow.com/questions/17546016/how-can-you-zip-or-unzip-from-the-command-prompt-using-only-windows-built-in-ca
				{powershell.exe -nologo -noprofile -command "Add-Type -A System.IO.Compression.FileSystem;[IO.Compression.ZipFile]::CreateFromDirectory('tmp','}
				to-local-file clean-path file
				{')"}
			]
		]][[
			rejoin [{zip -r "} to-local-file clean-path file {" tmp}]
		]])

		delete-dir path

		file
	]
]
