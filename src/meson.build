common_files = '''
	charsets.c
	substmap.c
	fileutil.c
	confutil.c
	numutils.c
	ole.c
'''.split()

executable('catdoc',
	[
		'catdoc.c',
		'reader.c',
		'writer.c',
		'analyze.c',
		'rtfread.c',
		common_files
	],
	include_directories : includes,
	install : true
)

executable('xls2csv',
	[
		'xls2csv.c',
		'sheet.c',
		'xlsparse.c',
		common_files
	],
	include_directories : includes,
	install : true
)

executable('catppt',
	[
		'catppt.c',
		'pptparse.c',
		common_files
	],
	include_directories : includes,
	install : true
)

if get_option('build_wordview')
	charset_conf.set('wish_path', get_option('wish_path'))
	wordview_tcl = configure_file(
		input : 'wordview.tcl.in',
		output : 'wordview',
		configuration : charset_conf,
	)
	install_data(wordview_tcl,
		install_dir : get_option('bindir'),
		install_mode : 'rwxr-xr-x',
	)
endif
