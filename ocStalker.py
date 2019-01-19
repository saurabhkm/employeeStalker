import win32com.client
import win32api
import time
import pylab as plt

#StalkingMinutes = raw_input('Stalk for how long(minutes): ')

StalkingDuration = 10.0 #* StalkingMinutes

comm = win32com.client.Dispatch('Communicator.UIAutomation')

statusDictionary = {1:'Offline', 2:'Available' ,10:'Busy', 34:'Away', 14: 'BRB'}

class Employee():
	"""Employee class"""

	availableDuration = 0
	awayDuration = 0
	busyDuration = 0
	offlineDuration = 0

	def __init__(self, email):
		global contact
		contact = comm.GetContact(email, comm.MyServiceId)
		self.name = contact.FriendlyName
		self.email = email
		#self.currentStatus = statusDictionary[contact.Status]

	def getStatus(self):
		return statusDictionary[contact.Status]

	def printInfo(self):
		print self.email+'|'+(self.name)+'|' \
		+str(self.availableDuration)+'|'+str(self.busyDuration)+'|' \
		+str(self.awayDuration)+'|'+str(self.offlineDuration)

	def plotInfo(self):
		labels = 'Available', 'Busy', 'Away', 'Ofline'
		explode=(0, 0.05, 0, 0)
		fracs = [self.availableDuration/StalkingDuration, self.busyDuration/StalkingDuration, 
		self.awayDuration/StalkingDuration, self.offlineDuration/StalkingDuration]
		print fracs
		plt.pie(fracs, explode=explode, labels = labels,autopct='%1.1f%%', shadow=True, startangle=90)
		plt.title('Stalking results')
		plt.show()

def Stalk():
	if E1.getStatus() == 'Available':
		E1.availableDuration += 1
		print 'available'
	elif E1.getStatus() == 'Busy':
		E1.busyDuration += 1
		print 'busy'
	elif E1.getStatus() == 'Away':
		E1.awayDuration += 1
		print 'away'
	else:
		E1.offlineDuration += 1
		print 'offline'

listToStalk = []

inputChoice = None

while inputChoice != 'done':
	inputChoice = raw_input("Enter Choice")
	listToStalk.append(inputChoice)

listToStalk.pop()
print 'stalking' + listToStalk + '... ;-)'

Elist = []

for employee in listToStalk:
	Elist = Elist.append(Employee(employee))

Stalk()

for k in range(int(StalkingDuration)-1):
	Stalk()
	time.sleep(1)

print E1.availableDuration, E1.awayDuration, E1.busyDuration
E1.printInfo()
E1.plotInfo()