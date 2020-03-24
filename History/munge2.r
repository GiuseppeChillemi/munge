Rebol [
	Title:		"Munge functions"
	Owner:		"Ashley G Truter"
	Version:	2.1.1
	Date:		22-Jun-2016
	Purpose:	"Extract and manipulate tabular values in blocks, delimited files and database tables."
	Licence:	"MIT. Free for both commercial and non-commercial use."
	Acknowledgments: {
		http://www.rebol.org/view-script.r?script=csv-tools.r
		http://www.rebol.org/view-script.r?script=flatten.r
		http://stackoverflow.com/questions/21716201/perform-file-encoding-conversion-with-rebol-3
		http://www.rebol.org/view-script.r?script=rebzip.r
		https://github.com/cyphre/r3-scripts/blob/master/arc/rebzip.r3
		http://stackoverflow.com/questions/31612164/does-anyone-have-an-efficient-r3-function-that-mimics-the-behaviour-of-find-any
		http://officeopenxml.com/anatomyofOOXML-xlsx.php
		http://www.adamkoch.com/2009/07/25/white-space-and-character-160/
	}
	Tested: {
		3.0.99.3.3	R3 64-bit	atronixengineering.com
	}
	Usage: {
		Word conventions
			b					block
			c					character
			f					file
			i					integer
			p					position
			s					string
			v					value

		Patches
			glob				used by like
			latin1-to-utf8		latin1 to UTF-8 support
			like				substitute for find/any
			read-string			latin1-to-utf8 and /lines fix
			remove-each			R3 compatibility with R2
			sort-all			R3 fix
			unique-skip			R3 fix

		Informational
			digit				charset [#"0" - #"9"]
			alpha				charset [#"A" - #"Z" #"a" - #"z"]
			alphanum			union alpha digit
			alphanums?			Returns TRUE if string only contains alphanums
			alphas?				Returns TRUE if string only contains alphas
			digits?				Returns TRUE if source not empty and only contains digits
			binary-file?		Returns TRUE if file is compressed, PDF or MS Office
			cols?				Number of columns in a delimited file or string
			fields?				Column names in a delimited file
			rows?				Number of rows in a delimited file or string
			sheets?				Sheet names
			spec?				Unique alphanumeric column words in a block

		General
			append-column		Append a column of values to a block
			remove-column		Remove a column of values from a block
			average-of			Average of values in a block
			flatten				Flatten a block
			max-of				Returns the largest value in a series
			min-of				Returns the smallest value in a series
			mixedcase			Converts string of characters to mixedcase
			put					Replaces the value following a key, and returns the new value
			sum-of				Sum of values in a block
			to-column-alpha		Convert numeric column reference into an alpha column reference
			to-column-number	Convert alpha column reference into a numeric column reference
			to-rebol-date		Converts a string date to a REBOL date
			to-string-date		Converts a string or REBOL date to a YYYY-MM-DD string
			to-string-time		Converts a string or REBOL time to a HH.MM.SS string

		Basic
			copy-row			Returns a block of values at the specified row in a block
			fetch				Retrieve block of values based on primary key
			index				Create a rowid index on a block
			make-map			Converts a block of values into key and value pairs
			pick-cell			Returns the value at the specified position in a block
			poke-cell			Changes a value at the given position
			remove-row			Deletes a row from a block

		Advanced
			load-dsv			Parses complex delimiter-separated values from a file or string
			load-excel			Loads an Excel file
			merge				Join outer block to inner block on keys
			munge				Load and/or manipulate a block of tabular (column and row) values
			read-pdf			Reads from a PDF file
			split-dsv			Parses simple delimiter-separated values from a file or string
			sqlcmd				Execute a SQL Server statement
			sqlite				Execute a SQLite statement
			unarchive			Decompress archive (only works with compression methods 'store and 'deflate)
			unzip				Uncompress file(s)
			write-excel			Write block of values to an Excel file
	}
]

;
;	Patches
;

glob: use [literal][
	literal: complement charset "*?"
	func [
		"Creates a PARSE rule matching VALUE expanding * (any characters) and ? (any one character)."
		value [any-string!] "Value to expand"
	][
		collect [
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
	]
]

;	latin1 to UTF-8 support

latin1-to-utf8: function [
	"Latin1 binary to UTF-8 string conversion."
	binary [binary!]
][
	mark: :binary
	while [mark: invalid-utf? mark][
		change/part mark to char! mark/1 1
	]
	replace/all trim/with deline to string! binary null to char! 160 #" "
]

;	substitute for find/any

like: function [
	"Finds a value in a series and returns the series at the start of it."
	series [any-string!] "Series to search"
	value [any-string! block!] "Value to find"
	/match
][
	either match [
		any [block? value value: glob value]
		parse series [some [result: value return (result)] fail]
	][
		skips: switch/default first value [#[none] #"*" #"?" ['skip]][reduce ['skip 'to first value]]
		any [block? value value: glob value]
		parse series [some [result: value return (result) | skips] fail]
	]
]

;	latin1-to-utf8 and /lines fix

read-string: function [
	"Read string(s) from a file."
	file [file!]
	/part "Partial read a given number of units (source relative)"
		length [integer!]
	/lines "Convert to block of strings"
][
	file: latin1-to-utf8 read file
	while [lf = last file][remove back tail file]
	all [
		lines
		remove-each line file: deline/lines file [empty? line]
	]
	also either part [copy/part file length] [file] file: none
]

;	remove-each - R3 compatibility with R2

if integer? remove-each i [][][
	*remove-each: :remove-each
	remove-each: function ['word data body][also data *remove-each :word data body]
]

;	sort/skip/all - R3 fix

sort-all: function [
	series [series!]
	size [integer!]
	/reverse "Reverse sort order"
] either [a 1 a 0] = sort/skip/all [a 1 a 0] 2 [[
	b: make block! rows: divide length? series size
	;	enblock
	repeat i rows [
		append/only b copy/part skip series i - 1 * size size
	]
	clear series
	;	flatten
	foreach row either reverse [sort/reverse b][sort b] [
		append series row
	]
	also series b: none
]][[
	either reverse [
		sort/skip/all/reverse series size
	][
		sort/skip/all series size
	]
]]

;	unique/skip - R3 fix

unique-skip: function [set [block!] size [integer!]][
	row: make block! size
	repeat i size [append row to-word ajoin ['c i]]
	skip?: unset!
	do compose/deep [
		remove-each [(row)] sort-all set size [
			either skip? = b: reduce [(row)] [true] [skip?: b false]
		]
	]
]

context [

	;
	;	Private
	;

	as-integer: func [
		"Converts a binary to an integer."
		value [binary!]
	][
		to integer! to string! value
	]

	base-path: system/script/path

	call-out: function [
		cmd [string!]
	][
		call/wait/output/error cmd stdout: make string! 65536 stderr: make string! 1024
		any [empty? stderr cause-error 'user 'message trim/lines stderr]
		replace/all stdout to char! 160 #" "
		deline stdout
	]

	delimiter?: function [
		s [string!]
	][
		all [not empty? s lf = last s remove back tail s]
		case [
			find s tab [tab]
			find s #"|" [#"|"]
			find s #"~" [#"~"]
			s [#","]
		]
	]

	first-line: function [
		data [file! string!]
	][
		first deline/lines either string? data [copy/part data 4096] [read-string/part data 4096]
	]

	get-ishort: func [
		"Converts a little-endian short to an integer."
		value [binary!]
	][
		to integer! reverse copy/part value 2
	]

	get-ilong: func [
		"Converts a little-endian long to an integer."
		value [binary!]
	][
		to integer! reverse copy/part value 4
	]

	line-count: function [
		source [string! binary!]
	][
		either empty? source [0][
			i: 1
			parse source [any [thru newline (++ i)]]
			i
		]
	]

	OLEDB: function [
		;	Excel, using https://www.microsoft.com/en-au/download/details.aspx?id=13255
		file [file!]
		cmd [string!]
		/hdr "First row contains columnnames not data"
	][
		any [exists? file cause-error 'access 'cannot-open file]
		all [find file %' cause-error 'user 'message "file name contains an invalid ' character"]
		s: call-out ajoin [
			{powershell -NoProfile -Command "}
			{$o=New-Object System.Data.OleDb.OleDbConnection('Provider=Microsoft.ACE.OLEDB.12.0;Data Source=''} to-local-file clean-path file {''}
			either %.accdb = suffix? file ["');"][ajoin [{;Extended Properties=''Excel 12.0 Xml;HDR=} either hdr "YES" "NO" {;IMEX=1''');}]]
			cmd
			either #";" = last cmd [
				ajoin [
					{$s.Connection=$o;}
					{$o.Open();}
					{$r=$s.ExecuteReader();}
					{$t=New-Object System.Data.DataTable;}
					{$t.Load($r);}
					{$r.Close();}
					{$o.Close();}
					{$t|ConvertTo-CSV -Delimiter '|' -NoTypeInformation}
				]
			][""]
			{"}
		]
		all [4 = length? find s lf clear s]	; empty? sheet
		also trim/tail trim s s: none
	]

	oledb?: function [
		"Returns true for an Excel or Access file."
		file [file!]
	][
		all [
			suffix? file
			binary-file? file
			find [%.xls %.xml %.acc] copy/part suffix? file 4
		]
	]

	oledb-cols?: function [
		file [file!]
		sheet [string! word!]
	][
		unless %.accdb = suffix? file [
			sheet: join sheet "$"
			all [find sheet " " insert sheet "''" append sheet "''"]
		]
		remove/part deline/lines OLEDB/hdr file ajoin [
			{$o.Open();$o.GetSchema('Columns')|where TABLE_NAME -eq '} sheet {'|sort ORDINAL_POSITION|select COLUMN_NAME;$o.Close()}
		] 2
	]

	oledb-rows?: function [
		file [file!]
		sheet [string! word!]
	][
		unless %.accdb = suffix? file [
			sheet: join sheet "$"
			all [find sheet " " insert sheet "''" append sheet "''"]
		]
		to-integer last parse OLEDB file ajoin [
			{$s=New-Object System.Data.OleDb.OleDbCommand('SELECT COUNT(*) FROM [} sheet {]');}
		] {"}
	]

	sql-cols?: function [
		data [string!]
	][
		line: first-line data
		all [tab = last line append line tab]
		length? parse/all line form tab
	]

	to-column-word: function [
		i [integer!]
	][
		to word! make string! reduce ['c i]
	]

	to-key: function [
		block [block!]
		width [integer!]
		columns [integer! block!]
	][
		rows: divide length? block width
		foreach offset to-block columns [
			loop rows [
				poke block offset either empty? val: pick block offset [0] [to integer! val]
				offset: offset + width
			]
		]
	]

	to-row-words: function [
		spec [integer! block!]
	][
		either integer? spec [
			row: make block! spec
			repeat i spec [append row to-column-word i]
		][
			row: make block! length? spec
			foreach i spec [append row to-column-word i]
		]
	]

	;
	;	Informational
	;

	set 'digit charset [#"0" - #"9"]
	set 'alpha charset [#"A" - #"Z" #"a" - #"z"]
	set 'alphanum union alpha digit

	set 'alphanums? function [
		"Returns TRUE if source only contains alphanums."
		source [string! binary!]
	][
		not find source complement alphanum
	]

	set 'alphas? function [
		"Returns TRUE if source only contains alphas."
		source [string! binary!]
	][
		not find source complement alpha
	]

	set 'digits? function [
		"Returns TRUE if source not empty and only contains digits."
		source [string! binary!]
	][
		all [not empty? source not find source complement digit]
	]

	set 'binary-file? function [
		"Returns TRUE if file is compressed, PDF or MS Office."
		;	https://en.wikipedia.org/wiki/List_of_file_signatures
		file [file!]
	][
		b: read/part file 8
		any [
			#{D0CF11E0A1B11AE1} = copy/part b 8		; office
			#{377ABCAF271C}		= copy/part b 6		; 7z
			#{526172211A07}		= copy/part b 6		; rar
			#{00010000}			= copy/part b 4		; accdb
			#{25504446}			= copy/part b 4		; pdf
			#{425A68}			= copy/part b 3		; bz2
			#{1F9D}				= copy/part b 2		; z
			#{1FA0}				= copy/part b 2		; z
			#{504B}				= copy/part b 2		; zip
			#{1F8B}				= copy/part b 2		; gz
		]
	]

	set 'cols? function [
		"Number of columns in a delimited file or string."
		data [file! string!]
		/sheet {Excel worksheet name (default is "Sheet 1")}
			name [string! word!]
	][
		length? either sheet [fields?/sheet data name] [fields? data]
	]

	set 'fields? function [
		"Column names in a delimited file or string."
		data [file! string!]
		/sheet {Excel worksheet name (default is "Sheet 1")}
			name [string! word!]
	][
		either any [string? data not oledb? data] [
			either line: first-line data [
				delimiter: delimiter? line
				all [delimiter = last line append line delimiter]
				foreach s b: parse/all line form delimiter [trim s]
				b
			][make block! 0]
		][
			oledb-cols? data any [name first sheets? data]
		]
	]

	set 'rows? function [
		"Number of rows in a delimited file or string."
		data [file! string!]
		/sheet {Excel worksheet name (default is "Sheet 1")}
		name [string! word!]
	][
		case [
			string? data [line-count data]
			oledb? data [oledb-rows? data any [name first sheets? data]]
			true [line-count read data]
		]
	]

	set 'sheets? function [
		"Sheet names."
		file [file!]
	][
		b: remove/part deline/lines OLEDB file {$o.Open();$o.GetSchema('Tables')|where TABLE_TYPE -eq "TABLE"|SELECT TABLE_NAME;$o.Close()} 2
		unless %.accdb = suffix? file [
			foreach s b [trim/with s "'"]
			remove-each s b [#"$" <> last s]
			foreach s b [remove back tail s]
		]
		also b b: none
	]

	set 'spec? function [
		"Unique alphanumeric column words in a block."
		fields [block!]
		/as-is "Do not strip non-alphanum or prepend with &"
	][
		b: make block! length? fields
		repeat i length? fields [
			field: form pick fields i
			unless as-is [
				remove-each c field [not find alphanum c]
				all [empty? field field: to-column-alpha i]
				insert field "&"
				while [find b to-word field][append field "*"]
			]
			append b to-word field
		]
		b
	]

	;
	;	General
	;

	set 'append-column function [
		"Append a column of values to a block."
		block [block!]
		size [integer!]
		value
		/dup "Duplicates the append a specified number of times"
			count [integer!]
	][
		any [
			integer? rows: divide length? block size
			cause-error 'user 'message "size not a multiple of length"
		]
		any [count count: 1]
		blk: make block! size + count * rows
		repeat row rows compose/deep [
			append blk copy/part skip block row - 1 * (size) (size)
			(
				either count = 1 [
					[append blk (value)]
				][
					[append/dup blk (value) (count)]
				]
			)
		]
		also append clear block blk blk: none
	]

	set 'remove-column function [
		"Remove a column of values from a block."
		block [block!]
		size [integer!]
		index [integer!]
	][
		any [
			integer? rows: divide length? block size
			cause-error 'user 'message "size not a multiple of length"
		]
		blk: make block! size - 1 * rows
		repeat row rows compose [
			append blk head remove at copy/part skip block row - 1 * (size) (size) (index)
		]
		also append clear block blk blk: none
	]

	set 'average-of function [
		"Average of values in a block."
		block [block!]
	][
		all [empty? block return none]
		total: 0
		foreach value block [total: total + value]
		total / length? block
	]

	set 'flatten function [
		"Flatten a block."
		block [block!]
	][
		result: copy []
		parse block rule: [
			any [
				pos: block! :pos into rule | skip (append/only result first pos)
			]
		]
		also result result: none
	]

	set 'max-of function [
		"Returns the largest value in a series."
		series [series!] "Series to search"
	][
		val: first series
		foreach v series [val: max val v]
	]

	set 'min-of function [
		"Returns the smallest value in a series."
		series [series!] "Series to search"
	][
		val: first series
		foreach v series [val: min val v]
	]

	set 'mixedcase function [
		"Converts string of characters to mixedcase."
		string [string!]
	][
		uppercase/part lowercase string 1
		foreach char [#"'" #" " #"-" #"." #","][
			all [find string char string: next find string char mixedcase string]
		]
		string: head string
	]

	unless value? 'put [ ; Red native
		set 'put function [
			"Replaces the value following a key, and returns the new value."
			map [map!]
			key
			value
		][
			append map reduce [key value]
			value
		]
	]

	set 'sum-of function [
		"Sum of values in a block."
		block [block!]
	][
		all [empty? block return none]
		total: 0
		foreach value block [total: total + value]
	]

	set 'to-column-alpha function [
		"Convert numeric column reference into an alpha column reference."
		number [integer!] "Column number between 1 and 702"
	][
		any [positive? number cause-error 'user 'message "Positive number expected"]
		any [number <= 702 cause-error 'user 'message "Number cannot exceed 702"]
		either number <= 26 [form #"@" + number][
			ajoin [
				#"@" + to-integer number - 1 / 26
				either zero? r: mod number 26 ["Z"][#"@" + r]
			] 
		]
	]

	set 'to-column-number function [
		"Convert alpha column reference into a numeric column reference."
		alpha [word! string! char!]
	][
		any [find [1 2] length? alpha: uppercase form alpha cause-error 'user 'message "One or two letters expected"]
		any [find charset [#"A" - #"Z"] last alpha cause-error 'user 'message "Valid characters are A-Z"]
		minor: subtract to-integer last alpha: uppercase form alpha 64
		either 1 = length? alpha [minor] [
			any [find charset [#"A" - #"Z"] first alpha cause-error 'user 'message "Valid characters are A-Z"]
			(26 * subtract to-integer first alpha 64) + minor
		]
	]

	set 'to-rebol-date function [
		"Converts a string date to a REBOL date."
		date [string!]
		/mdy "Month/Day/Year format"
		/ydm "Year/Day/Month format"
		/day "Day precededs date"
	][
		date: parse date "/-. "
		all [day remove date]
		to date! case [
			mdy		[ajoin [second date "-" first date "-" third date]]
			ydm		[ajoin [second date "-" third date "-" first date]]
			true	[ajoin [first date "-" second date "-" third date]]
		]
	]

	set 'to-string-date function [
		"Converts a string or REBOL date to a YYYY-MM-DD string."
		date [string! date!]
		/mdy "Month/Day/Year format"
		/ydm "Year/Day/Month format"
	][
		all [
			string? date
			date: to date! case [
				mdy		[date: parse date "/- " ajoin [second date "/" first date "/" third date]]
				ydm		[date: parse date "/- " ajoin [second date "/" third date "/" first date]]
				true	[date]
			]
		]
		ajoin [date/year "-" next form 100 + date/month "-" next form 100 + date/day]
	]

	set 'to-string-time function [
		"Converts a string or REBOL time to a HH.MM.SS string."
		time [string! date! time!]
	][
		all [
			string? time
			time: to time! trim/all copy time
		]
		ajoin [next form 100 + time/hour "." next form 100 + time/minute "." next form 100 + time/second]
	]

	;
	;	Basic
	;

	set 'copy-row function [
		"Returns a block of values at the specified row in a block." 
		block [block!]
		width [integer!]
		row [integer!]
	][
		copy/part skip block row - 1 * width width
	]

	set 'fetch function [
		"Retrieve block of values based on primary key."
		block [block!]
		size [integer!]
		key
		/nosort
	][
		any [
			integer? divide length? block size
			cause-error 'user 'message "size not a multiple of length"
		]
		any [nosort sort/skip block size]
		either start: index? find/skip block key size [
			end: index? find/skip/last block key size
			copy/part at block start end - start + size
		][
			make block! 0
		]
	]

	set 'index function [
		"Create a rowid index on a block."
		block [block!]
		width [integer!]
		part [integer! block!]
		/tight "Use ajoin if key is a block"
	][
		b: copy []
		either integer? part [
			repeat i divide length? block width [
				append b reduce [pick block part i]
				part: part + width
			]
		][
			row: to-row-words width
			width: to-row-words part
			do compose/deep [
				i: 1
				foreach [(row)] block [
					append b reduce [(either tight [[make string!]][[form]]) reduce [(width)] ++ i]
				]
			]
		]
		also to-map b b: row: width: none
	]

	set 'make-map function [
		"Converts a block of values into key and value pairs."
		block [block!]
		width [integer!]
		key [integer! block!]
		value [integer!]
		/tight "Use ajoin if key is a block"
		/key-type key-datatype [datatype!] "Datatype of key"
		/val-type val-datatype [datatype!] "Datatype of value"
	][
		b: copy []
		either integer? key [
			offset: value - key
			repeat i divide length? block width compose/deep [
				append b reduce [
					(either key-type [compose [to (key-datatype)]][]) pick block key
					(either val-type [compose [to (val-datatype)]][]) pick block key + offset
				]
				key: key + width
			]
		][
			row: to-row-words width
			width: to-row-words key
			do compose/deep [
				foreach [(row)] block [
					append b reduce [
						(either key-type [compose [to (key-datatype)]][]) (either tight [[make string!]][[form]]) reduce [(width)]
						(either val-type [compose [to (val-datatype)]][]) (to-column-word value)
					]
				]
			]
		]
		also to-map unique-skip b 2 b: row: width: none
	]

	set 'pick-cell function [
		"Returns the value at the specified position in a block."
		block [block!]
		width [integer!]
		position [pair!]
	][
		pick block position/y - 1 * width + position/x
	]

	set 'poke-cell function [
		"Changes a value at the given position."
		block [block!]
		width [integer!]
		position [pair!]
		value
	][
		poke block position/y - 1 * width + position/x value
	]

	set 'remove-row function [
		"Deletes a row from a block."
		block [block!]
		width [integer!]
		row [integer!]
	][
		remove/part skip block row - 1 * width width
	]

	;
	;	Advanced
	;

	set 'load-dsv function [
		"Parses complex delimiter-separated values from a file or string."
		source [file! string! binary!]
		/part "Offset position(s) to retrieve"
			columns [block! integer!]
		/with "Alternate delimiter (default is comma)"
			delimiter [char!]
		/blocks "Rows as blocks"
	][
		all [
			file? source
			case [
				oledb? source		[source: load-excel/string source 1]
				binary-file? source	[cause-error 'user 'message "Unsupported binary format"]
				true				[source: read-string source]
			]
		]
		all [binary? source source: latin1-to-utf8 source]
		all [empty? source return make block! 0]
		any [delimiter delimiter: delimiter? first-line source]

		either part [
			columns: to-block columns
			value: [
				{"} (clear v) x: to [{"} | end] y: (append/part v x y)
				any [{"} x: {"} to [{"} | end] y: (append/part v x y)]
				[{"} x: to [delimiter | crlf | cr | lf | end] y: (append/part v x y) | end]
				(all [find columns ++ col append b trim/lines copy v]) |
				copy x to [delimiter | crlf | cr | lf | end] (all [find columns ++ col append b trim/lines x])
			]
		][
			value: [
				{"} (clear v) x: to [{"} | end] y: (append/part v x y)
				any [{"} x: {"} to [{"} | end] y: (append/part v x y)]
				[{"} x: to [delimiter | crlf | cr | lf | end] y: (append/part v x y) | end]
				(append b trim/lines copy v) |
				copy x to [delimiter | crlf | cr | lf | end] (append b trim/lines x)
			]
		]

		v: make string! 32

		either blocks [
			cols: cols? source
			blk: copy []
			parse source [any [not end (col: 1 append/only blk b: make block! cols) value any [delimiter value] [crlf | cr | lf | end]]]
			b: blk
		][
			b: copy []
			parse source [any [not end (col: 1) value any [delimiter value] [crlf | cr | lf | end]]]
		]

		source: columns: value: x: y: v: cols: blk: col: none
		recycle
		b
	]

	set 'load-excel function [
		"Loads an Excel file."
		file [file!]
		sheet [string! word! integer!]
		/part "Offset position(s) / columns(s) to retrieve"
			columns [block! integer! string! word!]
		/where "Expression that can reference columns as F1, F2, etc"
			condition [string!]
		/string "Return string"
		/distinct
		/hdr "First row contains columnnames not data (not compatible with /part or /where)"
	][
		any [exists? file cause-error 'access 'cannot-open file]
		all [integer? sheet sheet: pick sheets? file sheet]
		all [find condition "'" condition: replace/all copy condition "'" "''"]
		either part [
			part: copy ""
			either integer? first columns: to-block columns [
				foreach i columns [append part ajoin ["F" i ","]]
			][
				foreach s fields: oledb-cols? file sheet [trim/all s]
				foreach s columns [
					s: trim/all form s
					either i: index? find fields s [
						append part ajoin ["F" i ","]
					][
						cause-error 'user 'message reform ["column" s "not found in sheet" sheet "of" to-local-file clean-path file]
					]
				]
			]
			remove back tail part
		] [part: "*"]

		unless %.accdb = suffix? file [
			sheet: join sheet "$"
			all [find sheet " " insert sheet "''" append sheet "''"]
		]

		cmd: ajoin [{$s=New-Object System.Data.OleDb.OleDbCommand('SELECT} either distinct { DISTINCT }{ } part { FROM [} sheet {]} either where reform [" WHERE" condition] "" {');}]

		s: either hdr [OLEDB/hdr file cmd][OLEDB file cmd]

		either empty? s [
			make either string string! block! 0
		][
			s: next find s lf
			also either string [s][load-dsv s] s: none
		]
	]

	set 'merge function [
		"Join outer block to inner block on keys." 
		block1 [block!] "Outer block"
		width1 [integer!]
		block2 [block!] "Inner block to index"
		width2 [integer!]
		cols [block!] "Offset position(s) to retrieve in merged block"
		keys [block!] "Outer/inner join column pairs"
		/default "Use spaces on inner block misses"
	][
		row1: to-row-words width1
		row2: copy []
		repeat i width2 [append row2 to-column-word width1 + i]

		if default [
			outer: copy []
			append/dup outer make string! 0 width2
		]

		part: to-row-words cols
	
		cols: copy []
		foreach [a b] keys [append cols to-column-word a]

		idx: index block2 width2 either 2 = length? keys [second keys] [munge/part keys 2 2]

		b: copy []
		either default [
			do has row2 compose/deep [
				foreach [(row1)] block1 [
					either i: select idx (either 2 = length? keys [to-column-word first keys] [compose/only [reform (cols)]]) [
						set [(row2)] skip block2 i - 1 * (width2)
					][
						set [(row2)] [(append/dup make block! width2 make string! 0 width2)]
					]
					append b reduce [(part)]
				]
			]
		][
			do has row2 compose/deep [
				foreach [(row1)] block1 [
					all [
						i: select idx (either 2 = length? keys [to-column-word first keys] [compose/only [reform (cols)]])
						set [(row2)] skip block2 i - 1 * (width2)
						append b reduce [(part)]
					]
				]
			]
		]

		also copy/deep b b: idx: cols: part: row1: row2: none
	]

	set 'munge function [
		"Load and/or manipulate a block of tabular (column and row) values."
		data [block! file! binary! vector!] "REBOL block, CSV or Excel file"
		spec [integer! block! none!] "Size of each record or block of heading words (none! gets cols? file)"
		/sheet {Excel worksheet (default is 1)}
			name [string! word! integer!]
		/update "Offset/value pairs (returns original block)"
			action [block!]
		/delete "Delete matching rows (returns original block)"
		/part "Offset position(s) to retrieve"
			columns [block! integer! word!]
		/where "Expression that can reference columns as c1, c2, etc"
			condition [block!]
		/compact "Remove blank rows"
		/only "Remove duplicate rows"
		/group "One of count, max, min or sum"
			having [word! block!] "Word or expression that can reference the initial result set column as count, max, etc"
		/order "Sort result set"
		/save "Write result to a delimited file"
			file [file!] "csv or txt"
		/list "Return new-line records"
	][
		either file? data [
			data: either oledb? data [
				any [name name: 1]
				all [integer? name name: pick sheets? data name]
				any [spec spec: length? oledb-cols? data name]
				load-excel data name
			][
				any [spec spec: cols? data]
				load-dsv data
			]
		][
			if binary? data [
				any [spec spec: cols? data]
				data: load-dsv data
			]
		]

		either integer? spec [
			row: to-row-words size: spec
		][size: length? row: copy spec]

		any [
			integer? rows: divide length? data size
			cause-error 'user 'message "size not a multiple of length"
		]

		case [
			update [
				b: copy []
				foreach [col val] action [
					all [word? col col: index? find row col]
					append b compose [
						poke data i + (col) (either all [word? val attempt [to datatype! val]] [compose [to (to datatype! val) pick data i + (col)]] [val])
					]
				]
				unless where [
					foreach col row [
						all [
							where: find flatten action col
							where: true
							break
						]
					]
				]
				i: 0
				either where [
					either condition [
						do compose/deep [
							foreach [(row)] data [all [(condition) (b)] i: i + (size)]
						]
					][
						do compose/deep [
							foreach [(row)] data [(b) i: i + (size)]
						]
					]
				][loop rows compose [(b) i: i + (size)]]
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
			any [part where] [
				columns: either part [
					all [block? columns 1 = length? columns columns: first columns]
					case [
						integer? columns [part: pick row columns]
						word? columns [part: columns]
						integer? first columns [
							part: copy []
							foreach i columns [append part pick row i]
							compose/only [reduce (part)]
						]
						true [
							part: copy columns
							compose/only [reduce (columns)]
						]
					]
				][compose/only [reduce (row)]]
				b: copy []
				do compose/deep [
					foreach [(row)] data [(
						either where [
							compose/deep [all [(condition) append b (columns)]]
						][
							compose [append b (columns)]
						]
					)]
				]
				all [
					part
					size: length? row: to-block part
				]
				data: b
			]
		]

		all [
			compact
			do compose/deep [remove-each [(row)] data [empty? form reduce [(row)]]]
		]

		all [
			only
			either size > 1 [unique-skip data size] [data: unique data]
		]

		if all [group not empty? data] [
			unless operation: any [
				all [find flatten to-block having 'avg 'avg]
				all [find flatten to-block having 'count 'count]
				all [find flatten to-block having 'max 'max]
				all [find flatten to-block having 'min 'min]
				all [find flatten to-block having 'sum 'sum]
			][cause-error 'user 'message "Invalid group operation"]
			case [
				operation = 'count [
					i: 0
					b: copy []
					either size = 1 [sort data][sort-all data size]
					group: copy/part data size
					loop divide length? data size compose/deep [
						either group = copy/part data (size) [++ i][
							insert insert tail b group i
							group: copy/part data (size)
							i: 1
						]
						data: skip data (size)
					]
					insert insert tail b group i
					data: b
					++ size
					append row operation
				]
				size = 1 [
					data: switch operation [
						avg [average-of data]
						max [max-of data]
						min [min-of data]
						sum [sum-of data]
					]
				]
				true [
					val: copy []
					b: copy []
					sort-all data size
					group: copy/part data size - 1
					loop divide length? data size compose/deep [
						either group = copy/part data (size - 1) [append val pick data (size)][
							insert insert tail b group (
								switch operation [
									avg	[[average-of val]]
									max [[max-of val]]
									min [[min-of val]]
									sum [[sum-of val]]
								]
							)
							group: copy/part data (size - 1)
							append val: copy [] pick data (size)
						]
						data: skip data (size)
					]
					insert insert tail b group switch operation [
						avg	[average-of val]
						max [max-of val]
						min [min-of val]
						sum [sum-of val]
					]
					data: b
					poke row size operation
				]
			]
			b: val: none
			all [block? having return munge/where data row having]
		]

		all [order sort-all data size]

		case [
			save [
				b: copy []
				i: 1
				loop divide length? data size compose/deep [
					s: copy ""
					loop (size) [
						insert insert tail s (
							either %.csv = suffix? file [
								[either any [find val: trim/with form pick data ++ i {"} "," find val lf] [ajoin [{"} val {"}]][val] ","]
							][
								[pick data ++ i "^-"]
							]
						)
					]
					append b head remove back tail s
				]
				write/lines file b
			]
			list [
				new-line/all/skip data true size
			]
		]

		also data data: b: s: none
	]

	set 'read-pdf function [	; requires pdftotext.exe from http://www.foolabs.com/xpdf/download.html
		"Reads from a PDF file."
		file [file!]
		/lines "Handles data as lines"
	][
		any [exists? file cause-error 'access 'cannot-open file]
		call/wait ajoin [{"} to-local-file join base-path %pdftotext.exe {" -nopgbrk -table "} to-local-file clean-path file {" $tmp.txt}]
		also either lines [
			remove-each line read-string/lines %$tmp.txt [empty? trim/tail line]
		][trim/tail read-string %$tmp.txt] delete %$tmp.txt
	]

	set 'split-dsv function [
		"Parses simple delimiter-separated values from a file or string."
		source [file! string!]
		/with "Alternate delimiter (default is tab, bar then comma)"
			delimiter [char!]
	][
		all [file? source source: read-string source]
		all [empty? source return make block! 0]
		any [delimiter delimiter: delimiter? first-line source]
		all [delimiter = last source append source: copy source delimiter]
		source: parse/all source join delimiter lf
		foreach s source [trim s]
		also source source: none
	]

	set 'sqlcmd function [
		"Execute a SQL Server statement."
		server [string!]
		database [string!]
		statement [string!]
		/key "Columns to convert to integer"
			columns [integer! block!]
		/headings "Keep column headings"
		/raw "Do not process return buffer"
		/affected "Return rows affected instead of empty block"
	][
		all ["SELECT" = copy/part trim copy statement 6 statement: ajoin ["SET ANSI_WARNINGS OFF" newline statement]]
		stdout: call-out reform compose ["sqlcmd -X -S" server "-d" database "-I -Q" ajoin [{"} statement {"}] {-W -w 65535 -s"^-"} (either headings [][{-h -1}])]

		all [raw return stdout]
		all [empty? stdout return either affected [make string! 0][make block! 0]]

		all [like first-line stdout "Msg*,*Level*,*State*,*Server" cause-error 'user 'message trim/lines find stdout "Line"]

		either "^/(" = copy/part stdout 2 [
			either affected [trim/with stdout "()^/"] [make block! 0]
		][
			stdout: copy/part stdout find stdout "^/^/("
			all [find ["" "NULL"] stdout return make block! 0]

			size: sql-cols? stdout

			all [tab = last stdout append stdout tab]
			stdout: parse/all stdout join tab lf
			all [headings remove/part skip stdout size size]

			any [integer? divide length? stdout size cause-error 'user 'message "size not a multiple of length"]

			foreach val stdout [
				all ["NULL" == trim val clear val]
			]

			all [
				key
				to-key skip stdout either headings [size][0] size columns
			]

			also stdout stdout: none
		]
	]

	set 'sqlite function [
		"Execute a SQLite statement."
		database [file!]
		statement [string!]
		/key "Columns to convert to integer"
			columns [integer! block!]
		/headings "Keep column headings"
		/raw "Do not process return buffer"
	][
		stdout: call-out reform compose [{sqlite3 -separator "^-"} (either headings ["-header"][]) to-local-file database ajoin [{"} trim/lines statement {"}]]

		all [raw return stdout]
		all [find ["" "^/"] stdout return make block! 0]

		also either key [
			size: sql-cols? stdout
			stdout: parse/all stdout join tab lf
			all [
				key
				to-key skip stdout either headings [size][0] size columns
			]
			stdout
		] [parse/all stdout join tab lf] stdout: none
	]

	set 'unarchive function [
		"Decompress archive (only works with compression methods 'store and 'deflate)."
		source [file! url!]
		/info "File names only"
	][
		blk: make block! 32
		parse read source [
			to #{504B0304}
			some [
				to #{504B0304} 4 skip
				2 skip ; version
				copy flags 2 skip (if not zero? flags/1 and 1 [return false])
				copy method 2 skip (method: either info [-1][get-ishort method])
				4 skip ; time & date
				copy crc 4 skip (crc: reverse crc) ; crc-32
				copy compressed-size 4 skip (compressed-size: get-ilong compressed-size)
				copy uncompressed-size-raw 4 skip (uncompressed-size: get-ilong uncompressed-size-raw)
				copy name-length 2 skip (name-length: get-ishort name-length)
				copy extrafield-length 2 skip (extrafield-length: get-ishort extrafield-length)
				copy name name-length skip (name: to-file name)
				extrafield-length skip
				data: compressed-size skip
				(
					uncompressed-data: switch/default method [
						0 [
							copy/part data compressed-size
						]
						8 [
							data: either zero? uncompressed-size [
								copy #{}
							][
								rejoin [#{789C} copy/part data compressed-size crc uncompressed-size-raw]
							]
							attempt [decompress/gzip data]
						]
					] [none]
					append blk name
					any [
						info
						append blk either all [#"/" = last name empty? uncompressed-data][none][uncompressed-data]
					]
				)
			]
			to end
		]
		blk
	]

	set 'unzip function [
		"Uncompress file(s)."
		file [file!]
		/folder "Create ZIP folder" 
		/only "Do not create sub-folders"
	][
		any [exists? file: clean-path file cause-error 'access 'cannot-open file]
		either folder [
			make-dir folder: replace copy file %.zip ""
		][
			folder: first split-path file
		]
		stdout: call-out ajoin [{"c:\Program Files\7-Zip\7z" } either only ["e"]["x"] { -y -o"} to-local-file folder {" "} to-local-file file {"}]
		files: copy []
		foreach line parse/all copy/part find stdout "Extracting" find stdout "Everything is Ok" "^/" [
			if "Extracting" = copy/part line 10 [
				trim remove/part line 10
				either only [
					either dir? file: join dirize folder last split-path to-rebol-file line [delete-dir file] [append files file]
				][
					append files join dirize folder to-rebol-file line
				]
			]
		]
		also files files: none
	]

	set 'write-excel function [
		"Write block(s) of values to an Excel file."
		file [file!]
		data [block!] "Name [string!] Data [block!] Width [integer! block!] records"
		/filter "Add auto filter"
	][
		any [%.xlsx = suffix? file cause-error 'user 'message "not a valid .xlsx file extension"]
		all [exists? path: %$tmp/ cause-error 'user 'message "$tmp already exists"]

		make-dir/deep join path %_rels
		make-dir/deep join path %xl/_rels
		make-dir/deep join path %xl/worksheets

		xml-content-types: copy ""
		xml-workbook: copy ""
		xml-workbook-rels: copy ""
		sheet-number: 1

		foreach [sheet-name block spec] data [
			unless empty? block [
				either block? spec [width: length? spec] [width: spec insert/dup spec: copy [] 10 width]
				unless integer? rows: divide length? block width [
					delete-dir path
					cause-error 'user 'message "width not a multiple of length"
				]

				append xml-content-types ajoin [{<Override PartName="/xl/worksheets/sheet} sheet-number {.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>}]
				append xml-workbook ajoin [{<sheet name="} sheet-name {" sheetId="} sheet-number {" r:id="rId} sheet-number {"/>}]
				append xml-workbook-rels ajoin [{<Relationship Id="rId} sheet-number {" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet} sheet-number {.xml"/>}]

				;	%xl/worksheets/sheet<n>.xml

				blk: ajoin [
					<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
					<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
						<cols>
				]
				repeat i width [
					append blk ajoin [{<col min="} i {" max="} i {" width="} pick spec i {"/>}]
				]
				append blk ajoin [</cols><sheetData>]
				i: 1
				loop rows [
					append blk <row>
					loop width [
						append blk case [
							number? val: pick block ++ i [
								ajoin [<c><v> val </v></c>]
							]
							#"=" = first val: form val [
								ajoin [<c><f> next val </f></c>]
							]
							true [
								foreach [char code][
									"&"		"&amp;"
									"<"		"&lt;"
									">"		"&gt;"
									"^/"	"&#10;"
									{"}		"&quot;"
									{'}		"&#39;"
								][replace/all val char code]
								ajoin [<c t="inlineStr"><is><t> val </t></is></c>]
							]
						]
					]
					append blk </row>
				]
				append blk </sheetData>
				all [filter append blk ajoin [{<autoFilter ref="A1:} to-column-alpha width rows {"/>}]]
				append blk </worksheet>
				write rejoin [path %xl/worksheets/sheet sheet-number %.xml] blk

				++ sheet-number
			]
		]

		write join path %"[Content_Types].xml" ajoin [
			<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
			<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
				<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
				<Default Extension="xml" ContentType="application/xml"/>
				<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
				xml-content-types
			</Types>
		]

		write join path %_rels/.rels ajoin [
			<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
			<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
				<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
			</Relationships>
		]

		write join path %xl/workbook.xml ajoin [
			<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
			<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x15" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">
				<workbookPr defaultThemeVersion="153222"/>
				<sheets>
					xml-workbook
				</sheets>
			</workbook>
		]

		write join path %xl/_rels/workbook.xml.rels ajoin [
			<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
			<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
				xml-workbook-rels
			</Relationships>
		]

		xml-content-types: xml-workbook: xml-workbook-rels: blk: none

		;	create xlsx file

		cmd: ajoin [
			{"c:\Program Files\7-Zip\7z" a "}
			to-local-file clean-path file {"}
		]
		change-dir path
		call-out cmd
		change-dir %..
		delete-dir path
		true
	]
]
