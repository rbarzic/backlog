test1:
	python pyxlsx.py

backlog:
	python ./backlog2.py --csv=november2015_1.csv --currentmonth=Dec15

backlog_jan2016:
	python ./backlog2.py --csv=december2015_1.csv --currentmonth=Jan16 --year2016
