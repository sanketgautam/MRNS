# Important Note: This is just a demo code, developed for POI (Proof Of Idea)
# For actual use, this script has lot much scope for optimization and improvements

# Mention: This script may not be functional as of now, as the result declaration website is changed

import urllib
import requests
import urllib.request
from bs4 import BeautifulSoup

url = "http://dolphintechnologies.in/manit/accessview.php"
#starting scholar number
sch = int(input('Enter Scholar Number to start with : '));
#ending scholar number
end_sch = int(input('Enter Last Scholar Number to be Fetched: '));
#semester
sem = int(input('Enter Semester: '));
#branch name (example MECH-1) - just for creating filename
branch_name = str(input('Enter Semester: '));
print("The results of Previous SEMESTER are :")
fo = open(branch_name+"_"+sem+"_semester.doc", "a")

while 1:
    params = urllib.parse.urlencode({'scholar': sch, 'semester': sem})
    binary_params = params.encode('utf-8')
    f = urllib.request.urlopen(url, binary_params)
    
    soup = BeautifulSoup(f)
    res=2;
    
    subss = soup.find_all("span",{"class":"style3"})
    numss = soup.find_all("span",{"class":"style3"})
    gradess = soup.find_all("span",{"class":"style3"})
    marks = soup.find_all("span",{"class":"style3"})
    links = soup.find_all("span",{"class":"style17"})
    results = soup.find_all("span",{"class":"style3"})
    a=0;b=0;c=0;s=0;k=0;g=0;i=0;

    print("\n");
    for mark in marks:
        b=b+1
        if b not in [1,2,3,4,6,8,9,10,11,12,13,14,16,17,18,21,22,15,19,20]:
            string = str(mark.text).strip(" ");
            string = string.replace("\n","");
            string = string.strip("\r");
            string = string.strip("\t");
            if b == 5:
                print("Name : ",end=" ");
                fo.write("Name :  ");
            elif b==7:
                print("Scholar : ",end=" ");
                fo.write("Scholar :  ");
            print("%s" %(string))
            fo.write(string);
            fo.write("\n");
        res = mark.text;
                
    print("\n");
    fo.write("\n");
    print("Subject       Obtained        Grade",end="");
    fo.write("Subject         Obtained         Grade \n");
    for subs in subss:
        s=s+1
        if s == 15:
            for nums in numss:
                k=k+1
                if k == 19:
                    for grades in gradess:
                        g=g+1
                        if g == 20:
                            
                            for sub,num,grade in zip(subs,nums,grades): 
                                m = str(sub);
                                n = str(num);
                                o = str(grade);
                                if(m != "<br/>"):

                                    string = m.replace(" ","");
                                    string = string.replace("\n","");
                                    #handling few error cases
                                    if(string == "MTH121" or string == "PHY113" or string == "HUM114" or string == "PHY118"):
                                        string = string[:-1];
                                    print("%s" %(string),end="\t\t")
                                    fo.write(string+"\t\t\t");

                                    string = n.replace(" ","");
                                    string = string.replace("\n","");
                                    print("%s" %(string),end="\t\t")
                                    fo.write(string+"\t\t\t");


                                    string = o.replace(" ","");
                                    string = string.replace("\n","");
                                    print("%s" %(string),end="\n")
                                    fo.write(string+"\n");


    for link in links:
        c=c+1
        if c in [2,3,4]:
            string = str(link.text).replace(" ","");
            string = string.replace("\n","");
            print("%s" %(link.text),end="\t\t")
            fo.write(link.text);
    fo.write("\t\t");
    string = str(res).replace(" ","");
    string = string.strip("\r");
    string = string.replace("\n","");
    st = "Result : "+string;
    
    print(st);
    print("\n");
    fo.write(st);
    fo.write("\n\n\n\n");
    
    if sch == end_sch:
        break;
    sch+=1;
fo.close();