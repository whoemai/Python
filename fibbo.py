n = int(input('valor: '))

a, b = 0, 1

while a < n:
    a, b = b, a + b

if (a == n):
    print('Este valor PERTENCE a sequencia.')
else:
    print('Este valor NÃO pertence a sequencia.')
