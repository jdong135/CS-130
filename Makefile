PYLINT_OPTS_SHEETS = --exit-zero --disable=C0103,C0116
PYLINT_OPTS_TESTS = --exit-zero --disable=C0103,C0116,R0904
# C0103: invalid-name does not comform to snake_case
# C0116: missing docstring
# R0904: too many public methods

.PHONY: format
format:
	autopep8 -i sheets/*.py
	autopep8 -i tests/*.py

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
	pylint --ignore=formulas.lark $(PYLINT_OPTS_SHEETS) sheets | tee logs/sheets_lint.txt
	pylint $(PYLINT_OPTS_TESTS) tests | tee logs/tests_lint.txt

.PHONY: clean
clean:
	pyclean .
	rm -f logs/*.stats
	rm -f test-outputs/*.json
