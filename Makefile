.PHONY: format
format:
	autopep8 -i sheets/*.py
	autopep8 -i tests/*.py
	
.PHONY: clean
clean:
	pyclean .