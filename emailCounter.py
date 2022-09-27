import win32com.client 
from datetime import datetime, timedelta

outlook = win32com.client.Dispatch('outlook.application')
inbox = outlook.GetNamespace('MAPI').GetDefaultFolder(6)
messages = inbox.Items

nowdate = datetime.now() 
lowerBound = nowdate.replace(hour=1, minute=30, second=0, microsecond=0)
upperBound = nowdate.replace(hour=23, minute=59, second=0, microsecond=0)


def countMessages(daysAgo):
  dumpList = []
  hashmap = {}

  for i in range(1, daysAgo):
    lowerBoundIter = lowerBound - timedelta(days=i)
    upperBoundIter = upperBound - timedelta(days=i)
    lowerBoundFilter = "' AND [ReceivedTime] >= '" + lowerBoundIter.strftime('%m/%d/%Y %H:%M %p') + "'"
    upperBoundFilter = "[ReceivedTime] < '"+upperBoundIter.strftime('%m/%d/%Y %H:%M %p')
    filter = upperBoundFilter + lowerBoundFilter


    for message in messages.Restrict(filter):      
      x = str(message.ReceivedTime)
      datepart = x.split(" ")[0]

      hashmap[datepart] = hashmap.get(datepart, 0) + 1
      dumpList.append(str(message))
    
  totalMessages = len(dumpList)

  print("Date\t\tCount\n")
  for i, k in hashmap.items():
    print(f"{i}\t{k}")

      
  return f"Total emails received: {totalMessages}"  

print(countMessages(20))