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

.PHONY: stresstest
stresstest:
	python3 tests/test_stresstest.py

.PHONY: clean
clean:
	pyclean .
	rm -f logs/*.stats