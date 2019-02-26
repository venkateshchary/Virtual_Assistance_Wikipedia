import wikipedia
from win32com.client import Dispatch


class VirtualAssistance:

    def __init__(self):
        '''
        connecting the windows sound API driver
        '''
        self.speak = Dispatch("SAPI.SpVoice")

    def arguments(self,):
        text_input = input("please enter the searching word: ")
        self.read_from_user(text_input)

    def read_from_user(self, userInput):
        print("entered input is :", userInput)
        return self.input_process(userInput)

    def input_process(self, userInput):
        text_input = ""
        search_result = wikipedia.search(userInput)
        if len(search_result) == 0:
            print("unable to find the search query please research:")
            self.arguments()
        if len(search_result) > 0:
            for i in search_result:
                text_input = text_input + self.input_summery(i)
        else:
            self.input_summery(search_result[0])
        self.speaking(text_input)

    def input_summery(self, text):
        summery = wikipedia.summary(text, sentences=2)
        return summery

    def speaking(self, text):
        self.speak.Speak(text)


if __name__ == "__main__":
    vm = VirtualAssistance()
    vm.arguments()
