import os

size = os.get_terminal_size()
print(size)


def drawGrid(row, col):
    if row > size[1] - 2:
        row = size[1] - 2

    if col > size[0] - 2:
        if (size[0] - 2) % 2 == 0:
            col = size[0] - 3

    for j in range(row):
        for i in range(col):
            if j == min(range(row)) and i == min(range(col)):
                print("┌", end="")
            elif j == min(range(row)) and i == max(range(col)):
                print("┐")
            elif j == max(range(row)) and i == min(range(col)):
                print("└", end="")
            elif j == max(range(row)) and i == max(range(col)):
                print("┘")
            elif j == min(range(row)) and i % 2 == 0:
                print("┬", end="")
            elif j == max(range(row)) and i % 2 == 0:
                print("┴", end="")
            elif i == min(range(col)) and j > 0:
                print("├", end="")
            elif i == max(range(col)) and j > 0:
                print("┤")
            elif i % 2 == 0:
                print("┼", end="")
            else:
                print("─", end="")
    print(row, col)


drawGrid(100, 700)
