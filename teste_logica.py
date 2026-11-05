from datetime import date
import logica

assert logica.limpar_telefone("+55 (32) 91234-5678") == "32912345678"
assert logica.telefone_valido("32912345678") and not logica.telefone_valido("1234")
assert logica.data_valida(29, 2) and not logica.data_valida(31, 2)
assert logica.proximo_aniversario(10, 11, date(2026, 11, 5)) == date(2026, 11, 10)
assert logica.proximo_aniversario(1, 11, date(2026, 11, 5)) == date(2027, 11, 1)
assert logica.proximo_aniversario(29, 2, date(2026, 11, 5)) == date(2027, 2, 28)
assert logica.classificar(30) == "URGENTE (30 dias)"
assert logica.classificar(31) == "ATENÇÃO (60 dias)"
assert logica.classificar(61) == "Fora da janela"
print("Todos os testes passaram.")
