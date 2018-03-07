import sys
def start_logging(filename):
    # LOGGING-----------------------------------
    te = open(filename,'w')  # File where you need to keep the logs


    class Unbuffered:
        def __init__(self, stream):
            self.stream = stream

        def write(self, data):
            self.stream.write(data)
            self.stream.flush()
            te.write(data)  # Write the data of stdout here to a text file as well


    sys.stdout = Unbuffered(sys.stdout)