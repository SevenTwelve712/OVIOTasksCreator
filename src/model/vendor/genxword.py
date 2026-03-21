"""Calculate the crossword and export image and text files."""

# Authors: David Whitlock <alovedalongthe@gmail.com>, Bryan Helmig
# Crossword generator that outputs the grid and clues as a pdf file and/or
# the grid in png/svg format with a text file containing the words and clues.
# Copyright (C) 2010-2011 Bryan Helmig
# Copyright (C) 2011-2020 David Whitlock
#
# Genxword is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# Genxword is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with genxword.  If not, see <http://www.gnu.org/licenses/gpl.html>.

import random, time, json
from operator import itemgetter
from collections import defaultdict
from pathlib import Path


from PIL import Image, ImageDraw

from Configs import PathConfig
from src.model.vendor.complexstring import ComplexString


class Crossword(object):
    def __init__(self, rows, cols, empty=' ', available_words=[]):
        self.rows = rows
        self.cols = cols
        self.empty = empty
        self.available_words = available_words
        self.let_coords = defaultdict(list)

    def prep_grid_words(self):
        self.current_wordlist = []
        self.let_coords.clear()
        self.grid = [[self.empty]*self.cols for i in range(self.rows)]
        self.available_words = [word[:2] for word in self.available_words]
        self.first_word(self.available_words[0])

    def compute_crossword(self, time_permitted=1.00):
        self.best_wordlist = []
        wordlist_length = len(self.available_words)
        time_permitted = float(time_permitted)
        start_full = float(time.time())
        while (float(time.time()) - start_full) < time_permitted:
            self.prep_grid_words()
            [self.add_words(word) for i in range(2) for word in self.available_words
             if word not in self.current_wordlist]
            if len(self.current_wordlist) > len(self.best_wordlist):
                self.best_wordlist = list(self.current_wordlist)
                self.best_grid = list(self.grid)
            if len(self.best_wordlist) == wordlist_length:
                break
        #answer = '\n'.join([''.join(['{} '.format(c) for c in self.best_grid[r]]) for r in range(self.rows)])
        answer = '\n'.join([''.join([u'{} '.format(c) for c in self.best_grid[r]])
                            for r in range(self.rows)])
        return answer + '\n\n' + str(len(self.best_wordlist)) + ' out of ' + str(wordlist_length)

    def get_coords(self, word):
        """Return possible coordinates for each letter."""
        word_length = len(word[0])
        coordlist = []
        temp_list =  [(l, v) for l, letter in enumerate(word[0])
                      for k, v in self.let_coords.items() if k == letter]
        for coord in temp_list:
            letc = coord[0]
            for item in coord[1]:
                (rowc, colc, vertc) = item
                if vertc:
                    if colc - letc >= 0 and (colc - letc) + word_length <= self.cols:
                        row, col = (rowc, colc - letc)
                        score = self.check_score_horiz(word, row, col, word_length)
                        if score:
                            coordlist.append([rowc, colc - letc, 0, score])
                else:
                    if rowc - letc >= 0 and (rowc - letc) + word_length <= self.rows:
                        row, col = (rowc - letc, colc)
                        score = self.check_score_vert(word, row, col, word_length)
                        if score:
                            coordlist.append([rowc - letc, colc, 1, score])
        if coordlist:
            return max(coordlist, key=itemgetter(3))
        else:
            return

    def first_word(self, word):
        """Place the first word at a random position in the grid."""
        vertical = random.randrange(0, 2)
        if vertical:
            row = random.randrange(0, self.rows - len(word[0]))
            col = random.randrange(0, self.cols)
        else:
            row = random.randrange(0, self.rows)
            col = random.randrange(0, self.cols - len(word[0]))
        self.set_word(word, row, col, vertical)

    def add_words(self, word):
        """Add the rest of the words to the grid."""
        coordlist = self.get_coords(word)
        if not coordlist:
            return
        row, col, vertical = coordlist[0], coordlist[1], coordlist[2]
        self.set_word(word, row, col, vertical)

    def check_score_horiz(self, word, row, col, word_length, score=1):
        cell_occupied = self.cell_occupied
        if col and cell_occupied(row, col-1) or col + word_length != self.cols and cell_occupied(row, col + word_length):
            return 0
        for letter in word[0]:
            active_cell = self.grid[row][col]
            if active_cell == self.empty:
                if row + 1 != self.rows and cell_occupied(row+1, col) or row and cell_occupied(row-1, col):
                    return 0
            elif active_cell == letter:
                score += 1
            else:
                return 0
            col += 1
        return score

    def check_score_vert(self, word, row, col, word_length, score=1):
        cell_occupied = self.cell_occupied
        if row and cell_occupied(row-1, col) or row + word_length != self.rows and cell_occupied(row + word_length, col):
            return 0
        for letter in word[0]:
            active_cell = self.grid[row][col]
            if active_cell == self.empty:
                if col + 1 != self.cols and cell_occupied(row, col+1) or col and cell_occupied(row, col-1):
                    return 0
            elif active_cell == letter:
                score += 1
            else:
                return 0
            row += 1
        return score

    def set_word(self, word, row, col, vertical):
        """Put words on the grid and add them to the word list."""
        word.extend([row, col, vertical])
        self.current_wordlist.append(word)

        horizontal = not vertical
        for letter in word[0]:
            self.grid[row][col] = letter
            if (row, col, horizontal) not in self.let_coords[letter]:
                self.let_coords[letter].append((row, col, vertical))
            else:
                self.let_coords[letter].remove((row, col, horizontal))
            if vertical:
                row += 1
            else:
                col += 1

    def cell_occupied(self, row, col):
        cell = self.grid[row][col]
        if cell == self.empty:
            return False
        else:
            return True

    # def remove_blank_lines(self):
    #     for i, row in enumerate(self.best_grid):
    #         if all(elem == self.empty for elem in row):
    #             del self.best_grid[i]
    #             self.rows -= 1
    #
    #     for j in self.cols:
    #         if all(row[j] == self.empty for row in self.best_grid):
    #             for row in self.best_grid:
    #                 del row[j]
    #                 self.cols -= 1

    def _cell_must_draw_diag_border(self, coord: tuple[int, int]):
        res = {"lu": False, "ru": False, "ld": False, "rd": False} # left up, right up, left down, right down
        bg = self.best_grid
        x, y = coord
        if y == 0 or x == 0 or bg[y - 1][x - 1] == bg[y - 1][x] == bg[y][x - 1] == self.empty:
            res["lu"] = True

        if y == 0 or x == self.cols - 1 or bg[y - 1][x + 1] == bg[y][x + 1] == bg[y - 1][x] == self.empty:
            res["ru"] = True

        if y == self.rows - 1 or x == 0 or bg[y + 1][x - 1] == bg[y + 1][x] == bg[y][x - 1] == self.empty:
            res["ld"] = True

        if y == self.rows - 1 or x == self.cols - 1 or bg[y + 1][x + 1] == bg[y + 1][x] == bg[y][x + 1] == self.empty:
            res["rd"] = True
        return res

    def gen_img(self, cell_size: int, border_size: int):
        CELL_BORDER = 2
        BLACK = (0, 0, 0)
        GREY = (191, 191, 191)
        width, height = cell_size * self.cols + border_size * 2, cell_size * self.rows + border_size * 2

        img = Image.new("RGB", (width, height), "white")
        draw = ImageDraw.Draw(img)

        borders_inner = []
        borders_outer = []

        for i, row in enumerate(self.best_grid):
            for j, cell in enumerate(row):
                x = j * cell_size + border_size
                y = i * cell_size + border_size

                if cell != self.empty: # рисуем саму ячейку
                    if i == 0 or self.best_grid[i - 1][j] == self.empty: # верхняя
                        borders_inner.append([(x, y), (x + cell_size, y)])
                        borders_outer.append([(x, y - border_size), (x + cell_size, y)])

                    if i == self.rows - 1 or self.best_grid[i + 1][j] == self.empty: # нижняя
                        borders_outer.append([(x, y + cell_size), (x + cell_size, y + cell_size + border_size)])

                    if j == 0 or self.best_grid[i][j - 1] == self.empty: # левая
                        borders_outer.append([(x - border_size, y), (x, y + cell_size)])
                        borders_inner.append([(x, y), (x, y + cell_size)])

                    if j == self.cols - 1 or self.best_grid[i][j + 1] == self.empty: # правая
                        borders_outer.append([(x + cell_size, y), (x + cell_size + border_size, y + cell_size)])

                    # Добавляем диагональные границы
                    diag_borders = ([(x - border_size, y - border_size), (x, y)],
                                    [ (x + cell_size, y - border_size), (x + cell_size + border_size, y)],
                                    [(x - border_size, y + cell_size), (x, y + cell_size + border_size)],
                                    [ (x + cell_size, y + cell_size), (x + cell_size + border_size, y + cell_size + border_size)])
                    # left up, right up, left down, right down
                    for border, need_to_draw in zip(diag_borders, self._cell_must_draw_diag_border((j, i)).values()):
                        if need_to_draw:
                            borders_outer.append(border)

                    # нижнюю и правую границы ячейки рисуем в любом случае
                    borders_inner.append([(x, y + cell_size), (x + cell_size, y + cell_size)])
                    borders_inner.append([(x + cell_size, y), (x + cell_size, y + cell_size)])

        for coord in borders_outer:
            draw.rectangle(coord, fill=GREY)
        for coord in borders_inner:
            draw.line(coord, fill=BLACK, width=CELL_BORDER)

        return img

if __name__ == '__main__':
    ROWS = 10
    COLS = 10
    words = (
        "word",
        "slut",
        "blue",
        "boobies",
        "inheritance",
        "tango",
        "permitted"
    )

    words = [[ComplexString(line.upper()), line] for line in words]
    cross = Crossword(ROWS, COLS, available_words=words)
    cross.compute_crossword(0.01)
    print(*cross.best_grid, sep="\n")
    print(cross.best_wordlist)

    img = cross.gen_img(50, 10)
    img.save(Path(PathConfig.SAVE_DIR, "grid.jpg"))