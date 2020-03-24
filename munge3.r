; Red [] ; commented as 'red is undefined in REBOL/Core
Rebol [
	Title:		"Munge functions"
	Owner:		"Ashley G Truter"
	Version:	3.0.5
	Date:		26-Oct-2018
	Purpose:	"Extract and manipulate tabular values in blocks, delimited files and database tables."
	Licence:	"MIT. Free for both commercial and non-commercial use."
	Tested: {
		Windows
			REBOL/Core		2.7.8		rebol.com
			R3/64-bit		3.0.99		atronixengineering.com/downloads
			RED/32-bit		0.6.3		red-lang.org
		macOS
			REBOL/Core		2.7.8		rebol.com
			RED/32-bit		0.6.3		red-lang.org
	}
	Changes: {
		Removed:
			renc support
			read-pdf
			/lines from read-string (use deline/lines instead)
			/compact from load-dsv
			/preserve from load-dsv
		Added:
			enzero
			unarchive
			archive
		Fixed:
			split-line bug
			write-dsv bug
			load-dsv bug
			to-string-time now returns "HH:MM:SS"
			put returns value
			latin1-to-utf8 handles split strings correctly
			load-dsv/part/flat formats data correctly
		Enhanced
			Added /map to sqlcmd
			Added /flat to load-dsv, load-excel and load-fixed
			load-fixed now auto detects widths
			read-string with large files about 10-15x faster
			load-dsv about 3x faster
			rows? faster
			Added /flat to list
			write-excel now uses archive instead of 7z
	}
	Usage: {
		append-column		Append a column of values to a block.
		archive				Compress block of file and data pairs.
		ascii-file?			Returns TRUE if file is ASCII.
		average-of			Average of values in a block.
		call-oledb			Call OLEDB via PowerShell returning STDOUT.
		call-out			Call OS command returning STDOUT.
		check				Verify data structure.
		cols?				Number of columns in a delimited file or string.
		delimiter?			Probable delimiter, with priority given to comma, tab, bar, tilde then semi-colon.
		delta				Remove source rows that exist in target.
		dezero				Remove leading zeroes from string.
		digit				DIGIT is a bitset! value: make bitset! #{000000000000FFC0}
		digits?				Returns TRUE if data not empty and only contains digits.
		distinct			Remove duplicate and empty rows.
		enblock				Convert a block of values to a block of row blocks.
		enzero				Add leading zeroes to a string.
		export				Export words to global context.
		fields?				Column names in a delimited file or string.
		first-line			Returns the first non-empty line of a file.
		flatten				Flatten nested block(s).
		latin1-to-utf8		Latin1 binary to UTF-8 string conversion.
		letter				LETTER is a bitset! value: make bitset! #{00000000000000007FFFFFE07FFFFFE0}
		letters?			Returns TRUE if data only contains letters.
		like				Finds a value in a series, expanding * (any characters) and ? (any one character), and returns TRUE if found.
		list				Sets the new-line marker to end of block.
		load-dsv			Parses delimiter-separated values into row blocks.
		load-excel			Loads an Excel file.
		load-fixed			Loads fixed-width values from a file.
		map-source			Maps source columns to target columns.
		max-of				Returns the largest value in a series.
		merge				Join outer block to inner block on primary key.
		min-of				Returns the smallest value in a series.
		mixedcase			Converts string of characters to mixedcase.
		munge				Load and/or manipulate a block of tabular (column and row) values.
		oledb-file?			Returns true for an Excel or Access file.
		order				Orders rows by one or more column offsets.
		read-binary			Read bytes from a file.
		read-string			Read string from a text file.
		remove-column		Remove a column of values from a block.
		replace-deep		Replaces all occurences of a search value with the new value in a block or nested block.
		rows?				Number of rows in a delimited file or string.
		second-last/penult	Returns the second last value of a series.
		sheets?				Excel sheet names.
		split-line			Splits and returns line of text as a block.
		sqlcmd				Execute a SQL Server statement.
		sqlite				Execute a SQLite statement.
		sum-of				Sum of values in a block.
		to-column-alpha		Convert numeric column reference to an alpha column reference.
		to-column-number	Convert alpha column reference to a numeric column reference.
		to-rebol-date		Convert a string date to a Rebol date.
		to-rebol-time		Convert a string date/time to a Rebol time.
		to-string-date		Convert a string or Rebol date to a YYYY-MM-DD string.
		to-string-time		Convert a string or Rebol time to a HH:MM:SS string.
		unarchive			Decompresses archive (only works with compression methods 'store and 'deflate).
		write-dsv			Write block(s) of values to a delimited text file.
		write-excel			Write block(s) of values to an Excel file.
	}
]

case [
	not rebol [
		build: 		'red
		platform:	system/platform
		target:		either system/build/config/target = 'IA-32 ['x32] ['x64]

		foreach word [compress deline invalid-utf? join reform to-rebol-file] [
			all [value? word print [word "already defined!"]]
		]

		compress: function [
			data
			/gzip
		] [
			data
		]

		deline: function [
			string [any-string!]
			/lines "Return block of lines (works for LF, CR, CR-LF endings) (no modify)"
			"Converts string terminators to standard format, e.g. CRLF to LF"
		] [
			trim/with string cr
			either lines [split string lf] [string]
		]

		invalid-utf?: function [
			"Checks UTF encoding; if correct, returns none else position of error"
			binary [binary!]
		] compose [
			find binary (make bitset! [192 193 245 - 255])
		]

		join: function [
			"Concatenates values."
			value "Base value"
			rest "Value or block of values"
		] [
			value: either series? value [copy value] [form value]
			repend value rest
		]

		reform: function [
			"Forms a reduced block and returns a string"
			value "Value to reduce and form"
		] [
			form reduce value
		]

		to-rebol-file: :to-red-file
	]
	system/product = 'atronix-view [
		build:		'r3
		platform:	system/platform/1
		target:		'x64

		put: function [
			"Replaces the value following a key, and returns the map"
			map [map!]
			key
			value
		] [
			append map reduce [key value]
			value
		]
	]
	2 = system/version/1 [
		build: 		'r2
		platform:	any [pick [none macOS Windows] system/version/4 'Linux]
		target: 	'x32

		all [platform = 'Windows call/show ""]

		function:	:funct
		map!:		:block!

		put: function [
			"Replaces the value following a key, and returns the map"
			map [block! hash!]
			key
			value
		] [
			either hash? map [
				append map key
			] [
				remove/part find/skip map key 2 2
				append map reduce [key value]
			]
			value
		]

		split: function [
			"Parses delimiter-separated values into a block"
			string [series!]
			delimiter [char!]
		] [
			all [empty? string return make block! 0]
			blk: collect [
				parse/all string [
					any [mk1: some [mk2: delimiter break | skip] (keep/only copy/part mk1 mk2)]
				]
			]
			all [delimiter = last string append blk copy ""]
			blk
		]
	]
	true [
		cause-error 'user 'message ["Unsupported Rebol version or derivative"]
	]
]

delete-dir: function [
	"Deletes a directory including all files and subdirectories"
	folder [file! url!]
] [
	if dir? folder [
		foreach file read folder [
			delete-dir rejoin [folder file]
		]
	]
	delete folder
]

ctx-munge: context [

	append-column: function [
		"Append a column of values to a block"
		block [block!]
		value
		/header heading
	] [
		unless empty? block [
			foreach row block [
				append row value
			]
			all [header poke block/1 length? block/1 heading]
		]
		block
	]

	archive: function [ ; https://en.wikipedia.org/wiki/Zip_(file_format) & http://www.rebol.org/view-script.r?script=rebzip.r
		"Compress block of file and data pairs"
		source [series!]
	] either build = 'r2 [[none]] [compose/deep [ ; R2 does not support compress/gzip or checksum/method data 'CRC32
		case [
			empty? source [none]
			not block? source [
				join #{1F8B08000000000002FF} next next head reverse/part skip tail compress/gzip source -8 4
			]
			true [
				to-short:	function [i] [copy/part reverse to binary! i 2]
				to-long:	function [i] [copy/part reverse to binary! i 4]

				bin: copy #{}
				dir: copy #{}

				foreach [file series] source [
					all [none? series series: make string! 0]

					any [file? file cause-error 'user 'message reform ["found" type? file "where file! expected"]]
					any [series? series cause-error 'user 'message reform ["found" type? series "where series! expected"]]

					method: either greater? length? series length? compressed-data: compress data: to binary! series [
						compressed-data: copy/part skip compressed-data 2 skip tail compressed-data -8
						#{0800}				; deflate
					] [
						compressed-data: data
						#{0000}				; store
					]

					offset: length? bin

					append bin rejoin [
						#{504B0304}			; Local file header signature
						#{1400}				; Version needed to extract (minimum)
						#{0000}				; General purpose bit flag
						method				; Compression method
						#{0000}				; File last modification time
						#{0000}				; File last modification date
						crc:				to-long (either build = 'red [[checksum]] [[checksum/method]]) data 'CRC32
						compressed-size:	to-long length? compressed-data
						uncompressed-size:	to-long length? data
						filename-length:	to-short length? file
						#{0000}				; Extra field length
						filename: to binary! file
						#{}					; Extra field
						compressed-data		; Data
					]

					append dir rejoin [
						#{504B0102}			; Central directory file header signature
						#{1400}				; Version made by
						#{1400}				; Version needed to extract (minimum)
						#{0000}				; General purpose bit flag
						method				; Compression method
						#{0000}				; File last modification time
						#{0000}				; File last modification date
						crc					; CRC-32
						compressed-size		; Compressed size
						uncompressed-size	; Uncompressed size
						filename-length		; File name length
						#{0000}				; Extra field length
						#{0000}				; File comment length
						#{0000}				; Disk number where file starts
						#{0000}				; Internal file attributes
						#{00000000}			; External file attributes
						to-long offset		; Relative offset of local file header
						filename			; File name
						#{}					; Extra field
						#{}					; File comment
					]
				]

				append bin rejoin [
					dir
					#{504B0506}			; End of central directory signature
					#{0000}				; Number of this disk
					#{0000}				; Disk where central directory starts
					entries: to-short divide length? source 2	; Number of central directory records on this disk
					entries				; Total number of central directory records
					to-long length? dir	; Size of central directory
					to-long length? bin	; Offset of start of central directory
					#{0000}				; Comment length
					#{}					; Comment
				]

				bin
			]
		]
	]]

	ascii-file?: function [ ; https://en.wikipedia.org/wiki/List_of_file_signatures
		"Returns TRUE if file is ASCII"
		source [file! url!]
	] [
		#{504B} <> read-binary/part source 2
	]

	average-of: function [
		"Average of values in a block"
		block [block!]
	] [
		all [empty? block return none]
		total: 0
		foreach value block [total: total + value]
		total / length? block
	]

	call-oledb: if platform = 'Windows [
		function [
			"Call OLEDB via PowerShell returning STDOUT"
			;	Excel, using https://www.microsoft.com/en-us/download/details.aspx?id=54920
			file [file! url!]
			cmd [string!]
			/hdr "First row contains columnnames not data"
		] compose/deep [
			any [exists? file cause-error 'access 'cannot-open reduce [file]]
			all [find file %' cause-error 'user 'message ["file name contains an invalid ' character"]]
			trim/tail trim call-out rejoin [
				(either target = 'x64 ["powershell "] ["C:\Windows\SysNative\WindowsPowerShell\v1.0\powershell.exe "])
				{-nologo -noprofile -command "}
				{$o=New-Object System.Data.OleDb.OleDbConnection('Provider=Microsoft.ACE.OLEDB.12.0;Data Source=''} to-local-file clean-path file {''}
				either %.accdb = suffix? file ["');"] [rejoin [{;Extended Properties=''Excel 12.0 Xml;HDR=} either hdr ["YES"] ["NO"] {;IMEX=1;Mode=Read''');}]]
				cmd
				either #";" = last cmd [
					rejoin [
						{$s.Connection=$o;}
						{$t=New-Object System.Data.DataTable;}
						{$t.Load($s.ExecuteReader());}
						{$o.Close();}
						{$t|ConvertTo-CSV -Delimiter `t -NoTypeInformation}
					]
				] [""]
				{"}
			]
		]
	]

	call-out: function [
		"Call OS command returning STDOUT"
		cmd [string!]
	] either build = 'r2 [[
		call/wait/output/error cmd stdout: make string! 65536 stderr: make string! 1024
		any [empty? stderr cause-error 'user 'message reduce [trim/lines stderr]]
		read-string to binary! stdout
	]] [[
		call/wait/output/error cmd stdout: make binary! 65536 stderr: make string! 1024
		any [empty? stderr cause-error 'user 'message reduce [trim/lines stderr]]
		read-string stdout
	]]

	check: function [
		"Verify data structure"
		data [block!]
	] [
		unless empty? data [
			cols: length? data/1
			i: 1
			foreach row data [
				any [block? row cause-error 'user 'message reduce [reform ["Row" i "expected block but found" type? row]]]
				all [zero? length? row cause-error 'user 'message reduce [reform ["Row" i "empty"]]]
				any [cols = length? row cause-error 'user 'message reduce [reform ["Row" i "expected" cols "column(s) but found" length? row]]]
				all [block? row/1 row cause-error 'user 'message reduce [reform ["Row" i "did not expect first column to be a block"]]]
				i: i + 1
			]
		]
		true
	]

	cols?: either platform = 'Windows [
		function [
			"Number of columns in a delimited file or string"
			data [file! url! binary! string!]
			/with
				delimiter [char!]
			/sheet {Excel worksheet name (default is "Sheet 1")}
				name [string! word!]
		] [
			length? case [
				sheet	[fields?/sheet data name]
				with	[fields?/with data delimiter]
				true	[fields? data]
			]
		]
	] [
		function [
			"Number of columns in a delimited file or string"
			data [file! url! binary! string!]
			/with
				delimiter [char!]
		] [
			length? either with [
				fields?/with data delimiter
			] [
				fields? data
			]
		]
	]

	delimiter?: function [
		"Probable delimiter, with priority given to comma, tab, bar, tilde then semi-colon"
		data [file! url! string!]
	] [
		data: first-line data
		counts: copy [0 0 0 0 0]
		foreach char data [
			switch char [
				#","	[counts/1: counts/1 + 1]
				#"^-"	[counts/2: counts/2 + 1]
				#"|"	[counts/3: counts/3 + 1]
				#"~"	[counts/4: counts/4 + 1]
				#";"	[counts/5: counts/5 + 1]
			]
		]
		pick [#"," #"^-" #"|" #"~" #";"] index? find counts max-of counts
	]

	delta: function [
		"Remove source rows that exist in target"
		source [block!]
		target [block!]
	] [
		unless any [empty? source empty? target] [
			remove-each row source [find/only target row]
		]
		source
	]

	dezero: function [
		"Remove leading zeroes from string"
		string [string!]
	] either build = 'r2 [[
		while [all [not empty? string #"0" = first string]] [remove string]
		string
	]] [[
		while [#"0" = first string] [remove string]
		string
	]]

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
	] [
		old-row: none
		remove-each row sort data [
			any [
				all [
					find ["" #[none]] row/1
					1 = length? unique row
				]
				either row == old-row [true] [old-row: row false]
			]
		]
		data
	]

	enblock: function [
		"Convert a block of values to a block of row blocks"
		data [block!]
		cols [integer!]
	] [
		all [block? data/1 return data]
		any [integer? rows: divide length? data cols cause-error 'user 'message ["Cols not a multiple of length"]]
		repeat i rows [
			change/part/only at data i copy/part at data i cols cols
		]
		data
	]

	enzero: function [
		"Add leading zeroes to a string"
		string [string!]
		length [integer!]
	] [
		insert/dup string #"0" length - length? string
		string
	]

	export: function [
		"Export words to global context"
		words [block!] "Words to export"
	] [
		foreach word words [
			do compose [(to-set-word word) (to-get-word in self word)]
		]
		words
	]

	fields?: either platform = 'Windows [
		function [
			"Column names in a delimited file or string"
			data [file! url! string!]
			/with
				delimiter [char!]
			/sheet {Excel worksheet name (default is "Sheet 1")}
				name [string! word!]
		] [
			either any [platform <> 'Windows string? data not oledb-file? data] [
				all [file? data not ascii-file? data cause-error 'user 'message ["Cannot use with binary file"]]
				either data: first-line data [
					split-line data any [delimiter delimiter? data]
				] [
					make block! 0
				]
			] [
				sheet: either sheet [name] [first sheets? data]
				unless %.accdb = suffix? data [
					sheet: rejoin [sheet "$"]
					if any [
						find sheet " "
						find sheet "-"
				 	] [
						insert sheet "''" append sheet "''"
					]
				]
				remove/part deline/lines call-oledb/hdr data rejoin [
					{$o.Open();$o.GetSchema('Columns')|where TABLE_NAME -eq '} sheet {'|sort ORDINAL_POSITION|select COLUMN_NAME;$o.Close()}
				] 2
			]
		]
	] [
		function [
			"Column names in a delimited file or string"
			data [file! url! string!]
			/with
				delimiter [char!]
		] [
			all [file? data not ascii-file? data cause-error 'user 'message ["Cannot use with binary file"]]
			either data: first-line data [
				split-line data any [:delimiter delimiter? data]
			] [
				make block! 0
			]
		]
	]

	first-line: function [
		"Returns the first non-empty line of a file"
		data [file! url! string!]
	] [
		foreach line deline/lines either file? data [
			latin1-to-utf8 read-binary/part data 4096
		] [
			copy/part data 4096
		] [
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
	] [
		all [empty? data return data]
		result: copy []
		foreach row data [
			append result row
		]
	]

	latin1-to-utf8: function [ ; http://stackoverflow.com/questions/21716201/perform-file-encoding-conversion-with-rebol-3
		"Latin1 binary to UTF-8 string conversion"
		data [binary!]
	] compose/deep [
		;	remove #"^@"
		trim/with data null
		;	remove #"^M" from split crlf
		either empty? data [make string! 0] [
			all [cr = last data take/last data]
			;	replace char 160 with space - http://www.adamkoch.com/2009/07/25/white-space-and-character-160/
			mark: data
			while [mark: find mark #{C2A0}] [
				change/part mark #{20} 2
			]
			;	replace latin1 with UTF
			(either build = 'r2 [] [[
				mark: data
				while [mark: invalid-utf? mark] [
					change/part mark to char! mark/1 1
				]
			]])
			deline to string! data
		]
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
		value [any-string!] "Value to find"
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
		/flat
			cols
	] [
		either all [flat cols > 1] [
			new-line/all/skip data true cols
		] [
			new-line/all data true
		]
	]

	load-dsv: function [ ; http://www.rebol.org/view-script.r?script=csv-tools.r
		"Parses delimiter-separated values into row blocks"
		source [file! url! binary! string!]
		/part "Offset position(s) to retrieve"
			columns [block! integer!]
		/where "Expression that can reference columns as row/1, row/2, etc"
			condition [block!]
		/with "Alternate delimiter (default is tab, bar then comma)"
			delimiter [char!]
		/flat "Return a block of values"
		/ignore "Ignore truncated row errors"
	] compose [
		source: case [
			file? source	[either ascii-file? source [read-string source] [cause-error 'user 'message ["Cannot use load-dsv with binary file"]]]
			binary? source	[read-string source]
			string? source	[deline source]
		]

		any [with delimiter: delimiter? source]

		(either build = 'r2 [[
			valchars: remove/part charset [#"^(00)" - #"^(FF)"] crlf
			valchars: compose [any (remove/part valchars form delimiter)]
			value: [
				{"} (clear v) x: [to {"} | to end] y: (insert/part tail v x y)
				any [{"} x: {"} [to {"} | to end] y: (insert/part tail v x y)]
				[{"} x: valchars y: (insert/part tail v x y) | end]
				(insert tail row trim/lines copy v) |
				x: valchars y: (insert tail row trim/lines copy/part x y)
			]
		]] [[
			value: [
				{"} (clear v) x: to [{"} | end] y: (append/part v x y)
				any [{"} x: {"} to [{"} | end] y: (append/part v x y)]
				[{"} x: to [delimiter | lf | end] y: (append/part v x y) | end]
				(append row trim/lines copy v) |
				copy x to [delimiter | lf | end] (append row trim/lines x)
			]
		]])

		v: make string! 32

		cols: cols?/with source delimiter

		append-row: function [
			row [block!]
		] compose/deep [
			all [
				row <> [""]
				(either where [condition] [])
				(either ignore [] [compose/deep [any [(cols) = len: length? row cause-error 'user 'message reduce [reform ["Expected" (cols) "values but found" len "on line" line]]]]])
				(either part [
					columns: to block! columns
					unless ignore [
						remove-each val integer-columns: copy columns [not integer? val]
						all [
							cols < i: max-of integer-columns
							cause-error 'user 'message reduce [reform ["Expected a value in column" i "but only found" cols "columns"]]
						]
					]
					part: copy/deep [reduce []]
					foreach col columns [
						append part/2 either integer? col [
							compose [(append to path! 'row col)]
						] [col]
					]
					compose [row: (part)]
				] [])
				(either flat [
					compose [row <> at tail blk (negate either part [length? columns] [cols]) append]
				] [
					compose [(either build = 'r2 [[row <> pick tail blk -1 append/only]] [[row <> last blk append/only]])]
				]) blk row
			]
		]

		line: 0

		blk: copy []

		(either build = 'r2 [[
			parse/all source [
				any [
					end break | (row: make block! cols)
					value
					any [delimiter value] [lf | end] (line: line + 1 append-row row)
				]
			]
		]] [[
			parse source [
				any [
					not end (row: make block! cols)
					value
					any [delimiter value] [lf | end] (line: line + 1 append-row row)
				]
			]
		]])

		either flat [
			list/flat blk either part [length? columns] [cols]
		] [list blk]
	]

	load-excel: if platform = 'Windows [
		function [
			"Loads an Excel file"
			file [file! url!]
			sheet [integer! word! string!]
			/part "Offset position(s) / columns(s) to retrieve"
				columns [block! integer!]
			/where "Expression that can reference columns as F1, F2, etc"
				condition [string!]
			/flat "Return a block of values"
			/distinct
			/string
		] [
			any [exists? file cause-error 'access 'cannot-open reduce [file]]
			all [where find condition "'" condition: replace/all copy condition "'" "''"]
			either part [
				part: copy ""
				foreach i to block! columns [append part rejoin ["F" i ","]]
				remove back tail part
			] [part: "*"]

			unless any [integer? sheet %.accdb = suffix? file] [
				sheet: rejoin [sheet "$"]
			]

			stdout: call-oledb file rejoin [
				{$o.Open();$s=New-Object System.Data.OleDb.OleDbCommand('}
				"SELECT " either distinct ["DISTINCT "] [""] part
				" FROM [" either integer? sheet [rejoin ["'+$o.GetSchema('Tables').rows[" sheet - 1 "].TABLE_NAME+'"]] [sheet] "]"
				either where [reform [" WHERE" condition]] [""]
				{');}
			]

			case [
				string	[stdout]
				flat	[remove/part load-dsv/flat/with stdout #"^-" cols? stdout]
				true	[remove load-dsv/with stdout #"^-"]
			]
		]
	]

	load-fixed: function [
		"Loads fixed-width values from a file"
		file [file! url!]
		/spec widths [block!]
		/part columns [integer! block!]
		/flat "Return a block of values"
	] [
		unless spec [
			widths: reduce [1 + length? line: first-line file]
			;	Red index? fails on none
			while [all [s: find/last/tail line "  " i: index? s]] [
				insert widths i
				line: trim copy/part line i - 1
			]

			insert widths 1

			repeat i -1 + length? widths [
				poke widths i widths/(i + 1) - widths/:i
			]

			take/last widths
		]

		spec: copy []
		pos: 1

		either part [
			part: copy []
			foreach width widths [
				append/only part reduce [pos width]
				pos: pos + width
			]
			foreach col to-block columns [
				append spec compose [trim copy/part at line (part/:col/1) (part/:col/2)]
			]
		] [
			foreach width widths [
				append spec compose [trim copy/part at line (pos) (width)]
				pos: pos + width
			]
		]

		blk: copy []

		foreach line deline/lines read-string file compose/deep [
			all [#"^L" = first line remove line]
			any [
				empty? trim copy line
				(either flat [[append]] [[append/only]]) blk reduce [(spec)]
			]
		]

		blk
	]

	map-source: function [
		"Maps source columns to target columns"
		source [file! url! string!]
		target [block!]
		map [map!]
		/verify "Number of column matches expected"
			columns [integer!]
		/load
	] [
		foreach field target [
			any [select map field put map field field]
			field2: trim/with copy field " _"
			any [select map field2 put map field2 field]
		]

		blk: copy []
		i: 1
		foreach field fields? source [
			append/only blk reduce [
				i
				to-column-alpha i
				field
				any [
					select map field
					select map trim/with copy field " _"
					select map i
					select map to-column-alpha i
				]
			]
			i: i + 1
		]

		all [
			verify
			columns <> len: length? munge/where blk [last row]
			cause-error 'user 'message reform ["Expected" columns "columns, found" len]
		]

		all [empty? blk return make block! 0]

		either load [
			part: flatten munge/part/where blk 1 [last row]
			source: either all [file? source oledb-file? source] [
				load-excel/part source 1 part
			] [
				load-dsv/part source part
			]
			any [1 < length? source return clear source]
			repeat i length? part [
				poke source/1 i last pick blk part/:i
			]
			source
		] [
			new-line/all blk true
		]
	]

	max-of: function [
		"Returns the largest value in a series"
		series [series!] "Series to search"
	] [
		all [empty? series return none]
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
	] [
		;	build rowid map of inner block
		map: make map! length? inner
		i: 0
		foreach row inner [
			put map row/:key2 i: i + 1
		]
		;	build column picker
		code: copy []
		foreach col columns [
			append code compose [(append to path! 'row col)]	; Ren-C fails as => b: copy [] b/1
		]
		;	iterate through outer block
		blk: make block! length? outer
		do compose/deep [
			either default [
				foreach row outer [
					all [
						(either build = 'r2 [[i: select/skip map row/:key1 2 i: first i]] [[i: select map row/:key1]])
						append row inner/:i
					]
					append/only blk reduce [(code)]
				]
			] [
				foreach row outer [
					all [
						(either build = 'r2 [[i: select/skip map row/:key1 2 i: first i]] [[i: select map row/:key1]])
						append row inner/:i
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
	] [
		all [empty? series return none]
		val: series/1
		foreach v series [val: min val v]
		val
	]

	mixedcase: function [
		"Converts string of characters to mixedcase"
		string [string!]
	] [
		uppercase/part lowercase string 1
		foreach char [#"'" #" " #"-" #"." #","] [
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
		/part "Offset position(s) and/or values to retrieve"
			columns
		/where "Expression that can reference columns as row/1, row/2, etc"
			condition
		/group "One of count, max, min or sum"
			having [word! block!] "Word or expression that can reference the initial result set column as count, max, etc"
	] [
		all [empty? data return data]

		if all [where not block? condition] [ ; http://www.rebol.org/view-script.r?script=binary-search.r
			lo: 1
			hi: rows: length? data
			mid: to integer! hi + lo / 2
			while [hi >= lo] [
				if condition = key: first data/:mid [
					lo: hi: mid
					while [all [lo > 1 condition = first data/(lo - 1)]] [lo: lo - 1]
					while [all [hi < rows condition = first data/(hi + 1)]] [hi: hi + 1]
					break
				]
				either condition > key [lo: mid + 1] [hi: mid - 1]
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
				either build <> 'red [
					either all [where block? condition] [
						foreach row data compose/deep [
							all [
								(condition)
								(blk)
							]
						]
					] [
						foreach row either where [copy/part at data lo rows] [data] compose [
							(blk)
						]
					]
				] [
					either block? condition [
						foreach row data bind compose/deep [
							all [
								(condition)
								(blk)
							]
						] 'row
					] [
						foreach row either where [copy/part at data lo rows] [data] bind compose [
							(blk)
						] 'row
					]
				]
				return data
			]
			delete [
				either where [
					either build <> 'red [
						remove-each row data compose/only [all (condition)]
					] [
						remove-each row data bind compose/only [all (condition)] 'row
					]
				] [
					clear data
				]
				return data
			]
			any [part where] [
				columns: either part [
					part: copy/deep [reduce []]
					foreach col to block! columns [
						append part/2 either integer? col [
							compose [(append to path! 'row col)]
						] [col]
					]
					part
				] ['row]
				blk: copy []
				foreach row data compose [
					(
						either where [
							either build <> 'red [
								compose/deep [all [(condition) append/only blk (columns)]]
							] [
								bind compose/deep [all [(condition) append/only blk (columns)]] 'row
							]
						] [
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
						either group = row [i: i + 1] [
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
						either group = copy/part row (len) [append val last row] [
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

		list data
	]

	oledb-file?: function [
		"Returns true for an Excel or Access file"
		file [file! url!]
	] [
		all [
			suffix? file
			find [%.acc %.xls] copy/part suffix? file 4
			exists? file
			not ascii-file? file
		]
	]

	order: function [
		"Orders rows by one or more column offsets"
		data [block!]
		columns [block! integer!] "Offset position(s) to order by"
		/desc "Descending order"
	] [
		any [1 < length? data return data]
		either integer? columns compose [
			(either desc [[sort/compare/reverse]] [[sort/compare]]) data function [a b] [a/:columns < b/:columns]
		] compose [
			by: compose [(append to path! 'a columns/1) < (append to path! 'b columns/1)]
			chain: compose [(append to path! 'a columns/1) = (append to path! 'b columns/1)]
			foreach i next columns [
				append by compose/only [all (copy chain)]
				append last by compose [(append to path! 'a i) < (append to path! 'b i)]
				append chain compose [(append to path! 'a i) = (append to path! 'b i)]
			]
			(either desc [[sort/compare/reverse]] [[sort/compare]]) data function [a b] compose/only [
				either any (by) [true] [false]
			]
		]
	]

	read-binary: function [
		"Read bytes from a file"
		source [file! url!]
		/part "Reads a specified number of bytes."
			length [integer!]
		/seek "Read from a specific position"
			index [integer!]
	] case [
		build = 'red [[
			case [
				all [part seek]	[read/binary/part/seek source length index]
				part			[read/binary/part source length]
				seek			[read/binary/seek source index]
				true			[read/binary source]
			]
		]]
		build = 'r3 [[
			case [
				all [part seek]	[read/part/seek source length index]
				part			[read/part source length]
				seek			[read/seek source index]
				true			[read source]
			]
		]]
		build = 'r2 [[
			case [
				all [part seek]	[skip read/binary/part source length + index index]
				part			[read/binary/part source length]
				seek			[skip read/binary source index]
				true			[read/binary source]
			]
		]]
	]

	read-string: function [
		"Read string from a text file"
		source [file! url! binary!]
	] compose [
		;	R2 read/skip does not work
		(either build = 'r2 [[any [binary? source source: read/binary source]]] [])
		i: 0
		either binary? source [
			s: make string! length: length? source
			while [i < length] [
				append s latin1-to-utf8 copy/part skip source i 1048576
				i: i + 1048576
			]
		] [
			s: make string! size: size? source
			while [i < size] [
				append s latin1-to-utf8 read-binary/seek/part source i 1048576
				i: i + 1048576
			]
		]
		s
	]

	remove-column: function [
		"Remove a column of values from a block"
		block [block!]
		index [integer!]
	] [
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
	] [
		repeat i length? data [
			either block? data/:i [replace-deep data/:i search new] [
				all [
					search = data/:i
					data/:i: new
				]
			]
		]
		data
	]

	rows?: function [
		"Number of rows in a delimited file or string"
		data [file! url! binary! string!]
	] [
		either any [
			all [file? data zero? size? data]
			empty? data
		] [0] [
			i: 1
			parse either file? data [read-binary data] [data] [
				any [thru newline (i: i + 1)]
			]
			i
		]
	]

	second-last: penult: function [
		"Returns the second last value of a series"
		string [series!]
	] [
		pick string subtract length? string 1
	]

	sheets?: if platform = 'Windows [
		function [
			"Excel sheet names"
			file [file! url!]
		] [
			blk: remove/part deline/lines call-oledb file {$o.Open();$o.GetSchema('Tables')|where TABLE_TYPE -eq "TABLE"|SELECT TABLE_NAME;$o.Close()} 2
			unless %.accdb = suffix? file [
				foreach s blk [trim/with s "'"]
				remove-each s blk [#"$" <> last s]
				foreach s blk [remove back tail s]
			]
			blk
		]
	]

	split-line: function [
		"Splits and returns line of text as a block"
		line [string!]
		delimiter [char!]
	] compose [
		all [empty? line return make block! 0]
		(
			either build = 'r2 [[
				valchars: remove/part charset [#"^(00)" - #"^(FF)"] crlf
				valchars: compose [any (remove/part valchars form delimiter)]
				value: [
					{"} (clear v) x: [to {"} | to end] y: (insert/part tail v x y)
					any [{"} x: {"} [to {"} | to end] y: (insert/part tail v x y)]
					[{"} x: valchars y: (insert/part tail v x y) | end]
					(insert tail b trim/lines copy v) |
					x: valchars y: (
						x: trim/lines copy/part x y
						all [
							not empty? x
							#"^"" = first x
							#"^"" = last x
							remove x
							take/last x
						]
						insert tail b x
					)
				]
			]] [[
				value: [
					{"} (clear v) x: to [{"} | end] y: (append/part v x y)
					any [{"} x: {"} to [{"} | end] y: (append/part v x y)]
					[{"} x: to [delimiter | lf | end] y: (append/part v x y) | end]
					(append b trim/lines copy v) |
					copy x to [delimiter | lf | end] (
						trim/lines x
						all [
							not empty? x
							#"^"" = first x
							#"^"" = last x
							remove x
							take/last x
						]
						append b x
					)
				]
			]]
		)
		v: make string! 32
		(
			either build = 'r2 [[
				parse/all line [
					any [
						end break | (b: copy [])
						value
						any [delimiter value] [lf | end] (return b)
					]
				]
			]] [[
				parse line [
					any [
						not end (b: copy [])
						value
						any [delimiter value] [lf | end] (return b)
					]
				]
			]]
		)
	]

	sqlcmd: if platform = 'Windows [
		function [
			"Execute a SQL Server statement"
			server [string!]
			database [string!]
			statement [string!]
			/key "Columns to convert to integer"
				columns [integer! block!]
			/headings "Keep column headings"
			/string
			/map "Returns 2 columns as map! (ignores /headings)"
			/affected "Return rows affected instead of empty block"
		] [
			stdout: call-out reform compose ["sqlcmd -X -S" server "-d" database "-I -Q" rejoin [{"} statement {"}] {-W -w 65535 -s"^-"} (either headings [] [{-h -1}])]

			all [string return stdout]
			all [#"^/" = first stdout return either affected [trim stdout] [make block! 0]]
			all [like first-line stdout "Msg*,*Level*,*State*,*Server" cause-error 'user 'message reduce [trim/lines find stdout "Line"]]

			stdout: copy/part stdout find stdout "^/^/("
			replace/all stdout "^/^/" "^/"
			all [find ["" "NULL"] stdout return make block! 0]

			either map [
				stdout: load-dsv/with/flat stdout #"^-"
				case [
					none? columns []
					block? columns [
						repeat i length? stdout [
							stdout/:i: to integer! stdout/:i
						]
					]
					columns = 1 [
						repeat i length? stdout [
							all [odd? i stdout/:i: to integer! stdout/:i]
						]
					]
					columns = 2 [
						repeat i length? stdout [
							all [even? i stdout/:i: to integer! stdout/:i]
						]
					]
				]
				return to map! stdout
			] [
				stdout: load-dsv/with stdout #"^-"
			]

			all [headings remove skip stdout 1]

			foreach row stdout [
				foreach val row [
					all ["NULL" == val clear val]
				]
			]

			remove-each row stdout [
				all [empty? row/1 1 = length? unique row]
			]

			all [
				key
				foreach row skip stdout either headings [1] [0] [
					foreach i to block! columns [
						row/:i: to integer! row/:i
					]
				]
			]

			stdout
		]
	]

	sqlite: function [
		"Execute a SQLite statement"
		database [file! url!]
		statement [string!]
		/key "Columns to convert to integer"
			columns [integer! block!]
		/headings "Keep column headings"
		/string
	] [
		stdout: call-out reform compose [{sqlite3 -separator "^-"} (either headings ["-header"] []) to-local-file database rejoin [{"} trim/lines statement {"}]]
		all [string return stdout]
		all [find ["" "^/"] stdout return make block! 0]
		stdout: load-dsv/with stdout #"^-"

		all [
			key
			foreach row skip stdout either headings [1] [0] [
				foreach i to block! columns [
					row/:i: to integer! row/:i
				]
			]
		]

		stdout
	]

	sum-of: function [
		"Sum of values in a block"
		block [block!]
	] [
		all [empty? block return none]
		total: 0
		foreach value block [total: total + value]
	]

	to-column-alpha: function [
		"Convert numeric column reference to an alpha column reference"
		number [integer!] "Column number between 1 and 702"
	] [
		any [positive? number cause-error 'user 'message ["Positive number expected"]]
		any [number <= 702 cause-error 'user 'message ["Number cannot exceed 702"]]
		either number <= 26 [form #"@" + number] [
			rejoin [
				#"@" + to integer! number - 1 / 26
				either zero? r: mod number 26 ["Z"] [#"@" + r]
			] 
		]
	]

	to-column-number: function [
		"Convert alpha column reference to a numeric column reference"
		alpha [word! string! char!]
	] [
		any [find [1 2] length? alpha: uppercase form alpha cause-error 'user 'message ["One or two letters expected"]]
		any [find letter last alpha cause-error 'user 'message ["Valid characters are A-Z"]]
		minor: subtract to integer! last alpha: uppercase form alpha 64
		either 1 = length? alpha [minor] [
			any [find letter alpha/1 cause-error 'user 'message ["Valid characters are A-Z"]]
			(26 * subtract to integer! alpha/1 64) + minor
		]
	]

	to-rebol-date: function [
		"Convert a string date to a Rebol date"
		date [string!]
		/mdy "Month/Day/Year format"
		/ydm "Year/Day/Month format"
		/day "Day precedes date"
	] compose/deep [
		all [day date: copy date remove/part date next find date " "]
		date: either digits? date [
			reduce [copy/part date 4 copy/part skip date 4 2 copy/part skip date 6 2]
		][
			(either build = 'red [[split date charset "/- "]] [[parse date "/- "]])
		]
		date/1: to integer! date/1
		date/2: to integer! date/2
		date/3: to integer! date/3
		to-date case [
			mdy		[reduce [date/2 date/1 date/3]]
			ydm		[reduce [date/2 date/3 date/1]]
			true	[reduce [date/1 date/2 date/3]]
		]
	]

	to-rebol-time: function [
		"Convert a string date/time to a Rebol time"
		time [string!]
	] either build <> 'red [[
		to time! trim/all copy back back find time ":"
	]] [[
		either find time "PM" [
			time: to time! time
			all [time/1 < 13 time/1: time/1 + 12]
			time
		] [
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
			date: case [
				mdy		[to-rebol-date/mdy date]
				ydm		[to-rebol-date/ydm date]
				true	[to-rebol-date date]
			]
		]
		all [
			date/year < 100
			date/year: date/year + either date/year <= (now/year - 1950) [2000] [1900]
		]
		rejoin [date/year "-" next form 100 + date/month "-" next form 100 + date/day]
	]

	to-string-time: function [
		"Convert a string or Rebol time to a HH:MM:SS string"
		time [string! date! time!]
	] compose [
		all [string? time digits? time time: rejoin [copy/part time 2 ":" copy/part skip time 2 2 ":" copy/part skip time 4 2]]
		(
			either build = 'red [[
				if string? time [
					PM?: find time "PM"
					time: to time! time
					all [PM? time/1 < 13 time/1: time/1 + 12]
				]
			]] [[
				all [
					string? time
					time: to time! trim/all copy time
				]
			]]
		)
		rejoin [next form 100 + time/hour ":" next form 100 + time/minute ":" next form 100 + to integer! time/second]
	]

	unarchive: function [ ; https://en.wikipedia.org/wiki/Zip_(file_format) & http://www.rebol.org/view-script.r?script=rebzip.r
		"Decompresses archive (only works with compression methods 'store and 'deflate)"
		source [file! url! binary!]
		/only file [file!]
		/info "File name/sizes only (size only for gzip)"
	] either build = 'r2 [[none]] [[ ; R2 does not support compress/gzip or checksum/method data 'CRC32
		any [binary? source source: read-binary source]

		case [
			#{1F8B08} = copy/part source 3 [
				either info [
					to integer! reverse skip tail copy source -4
				] [
					read-string either build = 'red [
						decompress source
					][
						decompress/gzip join #{789C} skip head reverse/part skip tail copy source -8 4 10
					]
				]
			]
			#{504B0304} <> copy/part source 4 [
				cause-error 'user 'message reform [source "is not a ZIP file"]
			]
			true [
				to-int: function [b] [to integer! reverse copy b]

				blk: make block! 32

				extract: either zero? source/7 and 1 [[
					;	Local file header - CRC-32, Compressed & Uncompressed fields precede data
					data: compressed-size skip
				]][[
					;	Data descriptor - data precedes CRC-32, Compressed & Uncompressed fields
					copy data to #{504B0708} 4 skip
					copy crc 4 skip (crc: reverse crc)
					copy compressed-size 4 skip (compressed-size: to-int compressed-size)
					copy size 4 skip (uncompressed-size: to-int size)
				]]

				parse source [
					some [
						#{504B0304} 4 skip
						copy method 2 skip
						4 skip
						copy crc 4 skip (crc: reverse crc)
						copy compressed-size 4 skip (compressed-size: to-int compressed-size)
						copy size 4 skip (uncompressed-size: to-int size)
						copy name-length 2 skip (name-length: to-int name-length)
						copy extrafield-length 2 skip (extrafield-length: to-int extrafield-length)
						copy name name-length skip (name: to-rebol-file to file! name)
						extrafield-length skip
						extract
						(
							append blk case [
								info				[reduce [name uncompressed-size]]
								#"/" = last name	[reduce [name none]]
								size = #{00000000}	[reduce [name make binary! 0]]
								method = #{0000}	[reduce [name copy/part data compressed-size]]
								true				[reduce [name either build = 'red [
														decompress/deflate rejoin [copy/part data compressed-size crc size] uncompressed-size
													][
														decompress/gzip rejoin [#{789C} copy/part data compressed-size crc size]]]
													]
							]
							all [only name = file return second blk]
						)
					]
					to end
				]

				either only [none] [blk]
			]
		]
	]]

	write-dsv: function [
		"Write block(s) of values to a delimited text file"
		file [file! url!] "csv or tab-delimited text file"
		data [block!]
	] [
		b: make block! length? data
		foreach row data compose/deep [
			s: copy ""
			foreach value row [
				append s (
					either %.csv = suffix? file [
						[rejoin [either any [find val: trim/with form value {"} "," find val lf] [rejoin [{"} val {"}]] [val] ","]]
					] [
						[rejoin [value "^-"]]
					]
				)
			]
			take/last s
			any [empty? s append b s]
		]
		write/lines file b
	]

	write-excel: function [ ; http://officeopenxml.com/anatomyofOOXML-xlsx.php
		"Write block(s) of values to an Excel file"
		file [file! url!]
		data [block!] "Name [string!] Data [block!] Widths [block!] records"
		/filter "Add auto filter"
	] either build = 'r2 [[none]] [[ ; archive function not available under R2
		any [%.xlsx = suffix? file cause-error 'user 'message ["not a valid .xlsx file extension"]]

		xml-content-types:	copy ""
		xml-workbook:		copy ""
		xml-workbook-rels:	copy ""
		xml-version:		{<?xml version="1.0" encoding="UTF-8" standalone="yes"?>}

		sheet-number:		1

		xml-archive:		copy []

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
					append blk rejoin [{<col min="} i {" max="} i {" width="} spec/:i {"/>}]
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
								foreach [char code] [
									"&"		"&amp;"
									"<"		"&lt;"
									">"		"&gt;"
									{"}		"&quot;"
									{'}		"&apos;"
									"^/"	"&#10;"
								] [replace/all value char code]
								rejoin [{<c t="inlineStr"><is><t>} value "</t></is></c>"]
							]
						]
					]
					append blk "</row>"
				]
				append blk {</sheetData>}
				all [filter append blk rejoin [{<autoFilter ref="A1:} to-column-alpha width length? block {"/>}]]
				append blk {</worksheet>}
				append xml-archive reduce [rejoin [%xl/worksheets/sheet sheet-number %.xml] blk]

				sheet-number: sheet-number + 1
			]
		]

		insert xml-archive reduce [
			%"[Content_Types].xml" rejoin [
				xml-version
				{<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
					<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
					<Default Extension="xml" ContentType="application/xml"/>
					<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>}
					xml-content-types
				{</Types>}
			]
			%_rels/.rels rejoin [
				xml-version
				{<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
					<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
				</Relationships>}
			]
			%xl/workbook.xml rejoin [
				xml-version
				{<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x15" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">
					<workbookPr defaultThemeVersion="153222"/>
					<sheets>}
						xml-workbook
					{</sheets>
				</workbook>}
			]
			%xl/_rels/workbook.xml.rels rejoin [
				xml-version
				{<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">}
					xml-workbook-rels
				{</Relationships>}
			]
		]

		write file archive xml-archive

		file
	]]
]
