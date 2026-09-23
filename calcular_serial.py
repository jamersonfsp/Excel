"""
Calculadora de Serial - NOC Analise de Viabilidade
=====================================================

Arquivo: NOC-Analise-de-Viabilidade-de-Economica-6.0.xlsm
Modulo:  EstaPastaDeTrabalho (Workbook_Open)
Linha:  CalculoSerial = GetDriveSerialNumber("C") * 2.54 / 3 + 15 * 10 / 17 * 1000

Como funciona:
1. Ao abrir o arquivo, o Excel mostra um MsgBox com o NUMERO (serial fisico
   do drive C: do usuario) -- e pede para o usuario informar ao desenvolvedor.
2. O desenvolvedor aplica a formula abaixo e devolve a SERIAL.
3. O usuario cola a SERIAL no InputBox. Se bater com
   Round(CalculoSerial, 0), a planilha abre. Senao, fecha.

USO:
    python3 calcular_serial.py <numero>
    python3 calcular_serial.py        # modo interativo

EXEMPLO:
    python3 calcular_serial.py 123456789
    >> Numero:  123456789
    >> Serial:  104520891
"""

import sys
import math


def calcular_serial(numero: int) -> int:
    """
    Replica exata da formula VBA:
        CalculoSerial = GetDriveSerialNumber("C") * 2.54 / 3 + 15 * 10 / 17 * 1000
        Serial = Round(CalculoSerial, 0)

    Atencao a precedencia de operadores do VBA (mesma do Excel):
        1. * e / tem mesma precedencia, avaliados da esquerda para direita
        2. + depois
    """
    # 15 * 10 / 17 * 1000 da esquerda pra direita:
    #   15 * 10 = 150
    #   150 / 17 = 8.82352941176...
    #   8.82352941176 * 1000 = 8823.52941176...
    termo_b = 15 * 10 / 17 * 1000

    # numero * 2.54 / 3 da esquerda pra direita:
    #   numero * 2.54
    #   resultado / 3
    termo_a = numero * 2.54 / 3

    calculo = termo_a + termo_b

    # VBA Round usa Banker's Rounding (arredondamento para o par mais proximo).
    # Para .5 exato, arredonda para o par (ex: 2.5 -> 2, 3.5 -> 4).
    # Python round() tambem faz Banker's Rounding, entao bate certinho.
    serial = round(calculo)
    return serial


def main():
    if len(sys.argv) > 1:
        try:
            numero = int(sys.argv[1].strip())
        except ValueError:
            print(f"Erro: '{sys.argv[1]}' nao e um numero inteiro valido.")
            sys.exit(1)
    else:
        print("=== Calculadora de Serial - NOC Viabilidade ===")
        print("Digite o NUMERO que aparece no MsgBox ao abrir a planilha.")
        print("(Ctrl+C para sair)\n")
        try:
            entrada = input("Numero: ").strip()
            numero = int(entrada)
        except KeyboardInterrupt:
            print("\nSaindo.")
            return
        except ValueError:
            print(f"Erro: '{entrada}' nao e um numero inteiro valido.")
            sys.exit(1)

    serial = calcular_serial(numero)

    print()
    print("=" * 50)
    print(f"  NUMERO (drive C): {numero}")
    print(f"  SERIAL (liberar): {serial}")
    print("=" * 50)
    print()
    print("Cole o SERIAL no InputBox da planilha para liberar o acesso.")
    print()
    # Mostra o calculo detalhado para conferencia
    termo_a = numero * 2.54 / 3
    termo_b = 15 * 10 / 17 * 1000
    calculo = termo_a + termo_b
    print("Conferencia (passo a passo):")
    print(f"  Termo A: {numero} * 2.54 / 3 = {termo_a:.6f}")
    print(f"  Termo B: 15 * 10 / 17 * 1000 = {termo_b:.6f}")
    print(f"  Soma: {calculo:.6f}")
    print(f"  Round(0): {serial}")


if __name__ == "__main__":
    main()
