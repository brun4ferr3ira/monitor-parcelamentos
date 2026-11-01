import csv
import unittest
from datetime import date
from pathlib import Path

import monitor_parcelamentos as mp

PASTA = Path(".")


class TestParcelasVencidas(unittest.TestCase):
    def test_exemplo_da_primeira_linha(self):
        primeira = date(2026, 3, 1)
        self.assertEqual(mp.parcelas_vencidas(primeira, 30, date(2026, 10, 31)), 8)

    def test_no_dia_do_vencimento_ainda_nao_venceu(self):
        primeira = date(2026, 3, 1)
        self.assertEqual(mp.parcelas_vencidas(primeira, 30, date(2026, 3, 1)), 0)
        self.assertEqual(mp.parcelas_vencidas(primeira, 30, date(2026, 3, 2)), 1)

    def test_nao_passa_da_quantidade_de_parcelas(self):
        self.assertEqual(mp.parcelas_vencidas(date(2026, 1, 10), 6, date(2026, 10, 31)), 6)


class TestClassificar(unittest.TestCase):
    limites = {"atencao": 1, "alerta": 2, "urgencia": 3, "rompimento": 4}

    def test_niveis(self):
        esperado = {0: "em_dia", 1: "atencao", 2: "alerta", 3: "urgencia", 4: "rompido", 7: "rompido"}
        for atrasadas, nivel in esperado.items():
            self.assertEqual(mp.classificar(10, atrasadas, self.limites), nivel)

    def test_quitado(self):
        self.assertEqual(mp.classificar(0, 0, self.limites), "quitado")


class TestGabarito(unittest.TestCase):
    def test_todos_os_parcelamentos_batem_com_o_gabarito(self):
        analise, _ = mp.carregar_analise(PASTA / "parcelamentos.xlsx", date(2026, 10, 31))
        calculado = {p["id"]: (p["atrasadas"], p["nivel"]) for p in analise}
        with open(PASTA / "gabarito.csv", encoding="utf-8") as f:
            for linha in csv.DictReader(f):
                esperado = (int(linha["parcelas_em_atraso"]), linha["nivel_esperado"])
                self.assertEqual(calculado[linha["id_parcelamento"]], esperado, linha["id_parcelamento"])
        self.assertEqual(len(calculado), 21)


if __name__ == "__main__":
    unittest.main()
