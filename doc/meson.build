docconf = configuration_data()
docconf.set('catdoc_version', meson.project_version())
docconf.set('charset_dir', charset_dir)

foreach f : [
	'catdoc.1',
	'catppt.1',
	'wordview.1',
	'xls2csv.1',
]
	man_page = configure_file(input: f + '.in',
		output : f,
		configuration : docconf
	)
	install_man(man_page)
endforeach
