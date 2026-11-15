# Testes automáticos simples das regras do sistema.
# Execute com:  python teste_logica.py   (se nenhum erro aparecer, está tudo certo)
from datetime import date
import logica

# Telefone: remove símbolos e o código do país (55)
assert logica.limpar_telefone("+55 (32) 91234-5678") == "32912345678"
# Telefone válido tem 10 ou 11 dígitos; "1234" é inválido
assert logica.telefone_valido("32912345678") and not logica.telefone_valido("1234")
# Data: 29/02 é aceito, 31/02 é bloqueado
assert logica.data_valida(29, 2) and not logica.data_valida(31, 2)
# Aniversário ainda neste ano (10/11 com hoje = 05/11/2026)
assert logica.proximo_aniversario(10, 11, date(2026, 11, 5)) == date(2026, 11, 10)
# Aniversário que já passou neste ano vai para o ano seguinte
assert logica.proximo_aniversario(1, 11, date(2026, 11, 5)) == date(2027, 11, 1)
# 29/02 em ano não bissexto vira 28/02
assert logica.proximo_aniversario(29, 2, date(2026, 11, 5)) == date(2027, 2, 28)
# Limites dos alertas: 30 = urgente, 31 = atenção, 61 = fora da janela
assert logica.classificar(30) == "URGENTE (30 dias)"
assert logica.classificar(31) == "ATENÇÃO (60 dias)"
assert logica.classificar(61) == "Fora da janela"
print("Todos os testes passaram.")