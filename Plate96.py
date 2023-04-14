import numpy as np


class Plate96:

    def __init__(self, cells, is_float):

        # General
        self.cells=cells
        self.is_float=is_float

        if self.is_float == 1:
            true_values=[x for x in cells if str(x) != 'nan']
            if true_values!=[]:
                self.mean = np.mean(true_values)
                self.cv = 100*np.std(true_values)/self.mean
                self.min = min(true_values)
                self.max = max(true_values)
            else:
                self.mean = None
                self.cv = None
                self.min = None
                self.max = None

        else:
            self.mean = None
            self.cv = None
            self.min = None
            self.max = None



if __name__ == '__main__':

    cells=[1,2,1,np.nan]
    test_plate=Plate96(cells,1)
    print(test_plate.mean)
    print(test_plate.cv)