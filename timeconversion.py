def convert(dat, position, mod):
    #mod is a modifier depending on whether we're looking at the end time or the start time 

    #it is used for dealing with converting to "military time"


    dashindex = str(dat.iloc[position][4]).index("-")

    #Find index of times
    def LocateTimes(dat):
        TimeArrays = []
        for i in range(len(dat)):
            if '-' in str(dat.iloc[i][4]):
                print(i)
                print(dat.iloc[i][4])
                TimeArrays.append(i)
        return TimeArrays

    TimeIndex=LocateTimes(dat)

    #Detect whether the time is am or pm based off user input 
    def appendampm(x):
        if x is None:
            return
        if 'am' in x.lower():
            y = 'am'
        elif 'pm' in x.lower():
            y = 'pm'
        else:
            return None, x
        return y, str(x)
    ampm, time = appendampm(str(dat.iloc[position][4][:dashindex]))

    #Extract Only Integers in Times
    def integerExtractor(var):
        integers = []
        inttemp = []
        #only get integers
        for x in var:
            try:
                x = int(x)
            except Exception as e:
                print(f'IntegerExtractor() returned an error): {str(e)}')

            else:
                integers.append(x)
        for a in integers:
            inttemp.append(str(a))
        output = ''.join(inttemp)
        integer = int(output)
        return integer
    #integer format
    time = integerExtractor(time)

    #Add colon for xx:xx if it is needed
    def colonorNot(data):
        temp=[]
        for x in data:
            temp.append(str(x))
        output = ''.join(temp)
        
        if len(output)==4:
            output = f"{output[:2]}:{output[2:]}"
        elif len(output)==3:
            output = f"{output[:1]}:{output[1:]}"
        else:
            return
        return output

    #Add zeros for time computation(1200, 1300, etc.)
    def zeros(data):
        if len(str(data))==1:
            slice = 1
            temp = list(str(data))
            temp.insert(slice, '00')
            data = int(''.join(temp))
            return data
        elif len(str(data))==2:
            slice = 2
            temp = list(str(data))
            temp.insert(slice, '00')
            data = int(''.join(temp))
            return data
        else:
            return(int(data))
    time=zeros(time)
    timeString = colonorNot(str(time))

    #Convert time to military time plus 2400 for easy computation

    def ManageAM(time, ampm, mod):
        if time == 1200 and ampm == 'am':
            time=0+mod
        elif ampm=='pm' and time<1200:
            time+=1200

            return time
    time = ManageAM(time, ampm, mod)

    return time, timeString


if __name__ == '__main__':
    convert()