from random import choice

f = float(input('Enter fundamental frequency:'))
num_of_ovtns = int(input('Enter number of overtones:'))
low = float(input('Enter low frequency:'))
high = float(input('Enter high frequency:'))
num_of_freqs = int(input('Enter number of frequencies to generate:'))
data = [ ]
for x in range(1, (num_of_ovtns+1)):
    s = f*x
    data.append(s)

file = open('spectrum.txt', 'w')
i = 0
while i < num_of_freqs:
    n = float(choice(data))
    if n < low:
        while n < low:
            n = n*2
            if n >= low:
                print(n)
                file.write(str(n)+'\n')
    else:
        if n > high:
            while n > high:
                n = n/2
                if n <= high:
                    print(n)
                    file.write(str(n)+'\n')
        else:
            print(n)
            file.write(str(n)+'\n')
    i += 1
file.close()

print('Done')
