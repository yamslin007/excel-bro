"""Independent in-memory mutants; never edit production/test files."""
import argparse
import subprocess
import sys


def run_case(case):
    import pytest
    from server.app.python_datasets import PythonDatasetRegistry, DatasetRejected
    attribute = "_grant" if case == "binding" else "valid" if case == "late-result" else "revoke"
    original = getattr(PythonDatasetRegistry, attribute)
    if case == "binding":
        def unbound(self, token, scope, origin):
            self._prune()
            if token not in self._grants:
                raise DatasetRejected()
            return self._grants[token]
        mutant = unbound
        selector = "exact_binding"
    elif case == "late-result":
        mutant = lambda *args, **kwargs: True
        selector = "late_result_discarded"
    else:
        mutant = lambda *args, **kwargs: None
        selector = "local_api_both_hosts_roundtrip"
    try:
        setattr(PythonDatasetRegistry, attribute, mutant)
        result = pytest.main(["server/tests/test_python_datasets.py", "-q", "-p", "no:cacheprovider", "-k", selector])
    finally:
        setattr(PythonDatasetRegistry, attribute, original)
    # 1 = actual assertion failures, not import/collection errors or no tests.
    if result != 1:
        raise SystemExit("RED-MISS: " + case)
    print("RED-OK: " + case)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--case", choices=["binding", "late-result", "revoke"])
    arguments = parser.parse_args()
    if arguments.case:
        run_case(arguments.case)
        return
    for case in ("binding", "late-result", "revoke"):
        subprocess.run([sys.executable, "-B", "-m", "scripts.validate_python_data_mutations", "--case", case], check=True)
        subprocess.run([sys.executable, "-B", "-m", "pytest", "server/tests/test_python_datasets.py", "-q", "-p", "no:cacheprovider"], check=True)
    print("All independent mutants RED; each fresh restored suite GREEN.")


if __name__ == "__main__":
    main()
