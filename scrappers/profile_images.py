#Edit value enclosed in angle brackets before executing the script

import urllib
import base64
start_sch = '<START_SCHOLAR_NUMBER>' #replace this with start value
end_sch = '<END_SCHOLAR_NUMBER>' #replace this with end value

#Note: start_sch & end_sch have same batch, course & branch

content=""

while True:
    #preparing image resource url
    url = 'http://manit.ecampuserp.com/assets/img/students/'+str(start_sch)+'.jpg'

    #fetching image data from url
    sock = urllib.urlopen(url)
    content = sock.read()
    sock.close()

    #saving image data as local file
    file = open("<PATH_TO_DEST_FOLDER>/"+str(start_sch)+".jpg",'wb+')
    file.write(content)
    file.close()

    print 'Image successfully saved for '+str(start_sch)
    print content
    start_sch+=1
    
    #breaking the loop after fetching all required images
    if(start_sch == end_sch)
        break;