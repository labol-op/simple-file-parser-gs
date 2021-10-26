# o framework de testes é o unittest, que é da biblioteca padrão
# https://docs.python.org/3/library/unittest.html

import unittest

class TestSimpleParser(unittest.TestCase):
    """
    Testes unitários para o parser.

    """
    def test_raise_import_error(self):
        """Testa se um erro é levantado quando importando um módulo que ainda não existe."""
        with self.assertRaises(ImportError):
            import simple_parser

if __name__ == "__main__":
    unittest.main()
