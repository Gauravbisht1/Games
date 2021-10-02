def speak(str):
    from win32com.client import Dispatch

    speak = Dispatch("SAPI.SpVoice")
    speak.Speak(str)

myboard={1:' ',2:' ',3:' ',4:' ',5:' ',6:' ',7:' ',8:' ',9:' '}
def printboard():

    print(f"{str(myboard.get(1,'+'))} | {str(myboard.get(2,'+'))} | {str(myboard.get(3,'+'))}")
    print("__|___|____")
    print(f"{str(myboard.get(4,'+'))} | {str(myboard.get(5,'+'))} | {str(myboard.get(6,'+'))}")
    print("__|___|__")
    print(f"{str(myboard.get(7,'+'))} | {str(myboard.get(8,'+'))} | {str(myboard.get(9,'+'))}")
    print("__|___|__")

def changes(num,inputs):
     myboard[num]=inputs

if __name__=="__main__":
 printboard()
 speak("Enter First player name")
 player1=input("Enter First player name --> ").title()
 speak("Enter Second player name")
 player2=input("Enter Second player name --> ").title()
 i=0
 import pyinputplus as hup
 win=True
 while i<10000:
  if i%2==0:
      speak(f"hey {player1},enter your number")
      print(player1,"---->", end=' ')
      try:
          num1 = hup.inputNum(min=1, lessThan = 10)
          if myboard[num1] == 'X' or myboard[num1] == 'O':
             speak("this spot is already filled, try again")
             print("this spot is already filled, try again\n")
             continue
      except:
          continue
      else:
          changes(num1,'X')
          i = i + 1
          printboard()

  elif i%2!=0:
      speak(f"hey {player2},enter your number")
      print(player2,"---->",end=' ')
      try:
          num2 = hup.inputNum(min=1,lessThan=10)
          if myboard[num2] == 'X' or myboard[num2] == 'O':
             speak("this spot is already filled, try again")
             print("this spot is already filled, try again\n")
             continue

      except:
          continue
      else:
          changes(num2,'O')
          i = i + 1
          printboard()



  if  (myboard[1] == myboard[2] == myboard[3] =='X' or myboard[1] == myboard[2] == myboard[3] == 'O' or
          myboard[4] == myboard[5] == myboard[6]=='X' or myboard[4] == myboard[5] == myboard[6] == 'O' or
            myboard[7] == myboard[8] == myboard[9]=='X' or myboard[7] == myboard[8] == myboard[9]=='O' or
            myboard[1] == myboard[4] == myboard[7]=='X' or myboard[1] == myboard[4] == myboard[7]=='O' or
            myboard[2] == myboard[5] == myboard[8]=='X' or myboard[2] == myboard[5] == myboard[8]=='O' or
            myboard[3] == myboard[6] == myboard[9]=='X' or myboard[3] == myboard[6] == myboard[9]=='O' or
            myboard[1] == myboard[5] == myboard[9]=='X' or myboard[1] == myboard[5] == myboard[9]=='O' or
            myboard[3] == myboard[5] == myboard[7]=='X' or myboard[3] == myboard[5] == myboard[7]=='O'):
        if (i+1)%2==0:
            print(f"{player1} won the game")
            speak(f"{player1} won the game")
        else:
            print(f"{player2} won the game")
            speak(f"{player2} won the game")
        break

  elif (' ' not in myboard.values() and
          ' ' not in myboard.values() and
          ' ' not in myboard.values() and
          ' ' not in myboard.values() and
          ' ' not in myboard.values() and
          ' ' not in myboard.values() and
          ' ' not in myboard.values() and
          ' ' not in myboard.values() and
          ' ' not in myboard.values()):
        print("Match Draw")
        speak("Match Draw")
        break
