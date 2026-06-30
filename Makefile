.PHONY: run test clean-data

run:
	python3 -m k12 --host 127.0.0.1 --port 8765

test:
	python3 -m unittest discover -s tests

clean-data:
	rm -rf .k12-data

