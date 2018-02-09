project('catdoc', 'c',
	default_options : [
		'c_std=c99',
		'buildtype=debugoptimized',
		'warning_level=1',
	],
	license : 'GPL2',
	version : '0.96',
	meson_version : '>= 0.44.0',
)

cc = meson.get_compiler('c')

add_project_arguments(
	'-include', 'config.h',
	language : 'c')

conf = configuration_data()
conf.set_quoted('CATDOC_VERSION', meson.project_version())
conf.set('_GNU_SOURCE', 1)

conf.set_quoted('SOURCE_CHARSET', get_option('input_charset'),
	description : 'Character encoding used by default for 8-bit source files')

conf.set_quoted('TARGET_CHARSET', get_option('output_charset'),
	description : 'Output character encoding used by default, if impossible to determine encoding from locale')

conf.set_quoted('UNKNOWN_CHAR', '?',
	description : 'Symbol to represent character which is not available either in target encoding or in replacement map')

if host_machine.endian() == 'big'
	conf.set('WORDS_BIGENDIAN', 1)
endif

charset_conf = configuration_data()
charset_dir = join_paths(get_option('prefix'), join_paths(get_option('libdir'), meson.project_name()))
charset_conf.set('charset_dir', charset_dir)

conf.set_quoted('CHARSETPATH', charset_dir)
conf.set_quoted('SYSTEMRC', join_paths(charset_dir, 'catdocrc'))

foreach h : [
	'stdint.h',
	'unistd.h',
]
	if cc.has_header(h)
		conf.set('HAVE_' + h.to_upper().underscorify(), 1,
			description : 'Define to 1 if you have the <' + h + '> header file')
	endif
endforeach

if get_option('use_langinfo')
	if cc.has_header('langinfo.h')
		conf.set('HAVE_LANGINFO', 1,
			description : 'Define to 1 if you have the <langinfo.h> header file')
	endif
endif

foreach f : [
	'strdup',
	'strftime',
	'strtol',
]
	if cc.has_function(f, args : '-D_GNU_SOURCE')
		conf.set('HAVE_' + f.to_upper(), 1,
			description : 'Define to 1 if you have the ' + f + ' function')
	endif
endforeach


config_h = configure_file(
	output : 'config.h',
	configuration : conf)

includes = include_directories('.')

subdir ('charsets')
subdir ('doc')
subdir ('src')
