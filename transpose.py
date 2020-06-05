from pynput.keyboard import Key, Controller #module used to emulate keypresses
import time #module used to give user time to switch to excel sheet after inputting info
k = Controller() #used to emulate keypresses
column= input("Enter column letter: ") #example: A, B, C, etc.
rowstart = int(input("Enter starting row number: ")) #example: 1, 2, 3, etc.
rowend = int(input("Enter ending row number: ")) #example: 1, 2, 3, etc.
interval = int(input("Enter blocksize: ")) #this is how many rows to transpose at once, examples: 1, 2, 3, etc.
r = 1 #this is here to keep track of the destination row
destcolumn = chr(ord(column)+2) #this is the destination column
print('Starting in 5 seconds [alt tab to your excel sheet]')
time.sleep(5) #give user time to switch back to excel
#for loop in range of some strange algorithm I figured out through some trial and error
for i in range(interval + rowstart - 1, rowend + 1, interval):
    #ctrl + g
    k.press(Key.ctrl)
    k.type('g')
    k.release(Key.ctrl)
    #enter in selection
    k.type(f'{column}{i}:{column}{i-interval+1}\n')
    #ctrl c
    k.press(Key.ctrl)
    k.type('c')
    k.release(Key.ctrl)
    #ctrl g
    k.press(Key.ctrl)
    k.type('g')
    k.release(Key.ctrl)
    #go to destination row and column
    k.type(f'{destcolumn}{r}\n')
    #transpose shortcut
    k.press(Key.alt)
    k.type('e')
    k.release(Key.alt)
    k.type('s')
    k.press(Key.alt)
    k.type('e')
    k.release(Key.alt)
    k.type('\n')
    #remove selection
    k.press(Key.esc)
    k.release(Key.esc)
    #for each transposition, the next one will occurr on the next row
    r += 1
#tell the user the program has finished executing
print("done")