.PHONY: format
format:
	autopep8 -i sheets/*.py
	autopep8 -i tests/*.py

.PHONY: test
test:
	python3 tests/test_github.py
	python3 tests/test_tarjan.py

.PHONY: clean
clean:
	pyclean .