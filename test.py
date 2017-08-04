
def test_func(num):


  print "State: " + state

  if state=="start":
    count=num
  else:
    count+=num

  print count

state="start"
test_func(10)
state="running"
test_func(10)




