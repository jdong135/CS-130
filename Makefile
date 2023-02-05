.PHONY: format
format:
	autopep8 -i sheets/*.py
	autopep8 -i tests/*.py

.PHONY: test
test:
	python3 tests/test_workbook.py
	python3 tests/test_lark_module.py
	python3 tests/test_spec1.py
	python3 tests/smoketest.py
	python3 tests/test_system.py
	python3 tests/test_spec2.py

stresstest:
	python3 tests/test_stresstest.py

.PHONY:
lint:
	pylint --ignore=formulas.lark --exit-zero sheets | tee logs/sheets_lint.txt
	pylint --exit-zero tests | tee logs/tests_lint.txt

.PHONY: clean
clean:
	pyclean .
	rm -f logs/*.stats
	rm -f test-outputs/*.json
