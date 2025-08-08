#
# reformat styles and and re-zip "example.odt"
#
test:
	cp example.unzipped/styles.xml .
	go run ./libreoffice_style_renamer.go styles.xml
	test -f styles_modified.xml

unzip:
	-mkdir example.unzipped
	cp ./example.odt ./example.odt.zip
	(cd example.unzipped; unzip ../example.odt.zip)
	rm ./example.odt.zip

zip:
	cd example.unzipped; \
	zip -0 ../example2.odt mimetype; \
	zip --recurse-paths ../example2.odt * --exclude mimetype

