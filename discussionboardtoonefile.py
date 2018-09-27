import requests
import csv

'''
This script makes a dump of the contents of the discussion board on Canvas and writes everything to one file. 
It requires the course_id in which the discussion board is present. You can find it 
by going to the course on canvas and look at the browser url the number after /courses/
is your course_id.
It also requires a valid Token which you can create on your personal settings page (Don't forget to put
an expiration date in).

NOTE: to get the contents including replies from the Canvas API I have used the entries option
in the documentation it states 

"Retrieve the (paginated) top-level entries in a discussion topic.

May require (depending on the topic) that the user has posted in the topic. If it is required, and the user has not posted, will respond with a 403 Forbidden status and the body 'require_initial_post'.

Will include the 10 most recent replies, if any, for each entry returned.

If the topic is a root topic with children corresponding to groups of a group assignment, entries from those subtopics for which the user belongs to the corresponding group will be returned.

Ordering of returned entries is newest-first by posting timestamp (reply activity is ignored)."

I have set a pagination to allow 100 pages however wasn't able to test on posts with more then 10 comments or 10 replies
some entries may be lost on bigger discussions

to excute this script scroll all the way down to the last line and replace class_id with the actually number inbetween ""
same counts for the token
if you do not put it between "" the script won't be able to read the class_id and token properly

NOTE NOTE: this script has been written for python 3.6.3 
'''

def discussion_content(class_id,token):
    print("Waking up")
    print("Need to clean systems")
    #Start url for UvA
    url = "https://canvas.uva.nl/api/v1/"
    #
    topic_url = url + "/courses" + "/" + class_id + "/discussion_topics"
    # Add token to header to authorize
    headers = {'Authorization': 'Bearer ' + str(token)}
    print("Grabbing the correct data")
    r = requests.get(topic_url, headers=headers, params={'per_page': 100}) #gets data from canvas per_page increase if more than 100 discussion topics are present
    print("got it", topic_url)
    r = r.json() #converst raw data to json format that python can read
    topic_list = [] #make list to at all topic numbers later
    print("all systems in order ready to start working")
    #finds all topic numbers and adds them to the topic_list as strings
    for i in r:
        x = str(i['id'])
        topic_list.append(x)
    print("found all discussion topics IDs ready to extract their info")
    #adding headers to csv
    f = csv.writer(open('discussions.csv', 'a'))
    f.writerow(['comment/reply', 'id', 'parent_id', 'created_at', 'updated_at', 'message', 'user_name',
                'discussion_topic_id'])
    #collects all data about the topics and writes the id, title, posted_at, last_reply_at message and user_name to topic_info.csv (appears in same folder as this script)
    print(topic_list, " writing topic info to file")
    for i in r:
        f = csv.writer(open("discussions.csv",'a'))
        f.writerow(['dicussion topic',
                    i['id'],
                    i['title'],
                    i['posted_at'],
                    i['last_reply_at'],
                    i['message'],
                    i['user_name'],
                    i['id']])

          # These are the headers for the csv
    #This part graps each row from topic_info.csv and writes that row (with the info of 1 topic) to a csv in which later the comments and replies will be added
    print("Extracted all info about discussion boards")

    '''
    for q in topic_list:
            with open("topic_info.csv",newline='') as File:
            reader = csv.reader(File)
            for row in reader:
                if q in row:
                    f = csv.writer(open('discussions.csv', 'a'))
                    f.writerow(row)
                else:
                    continue
    print("Starting to extract comments and replies of the discussion board")
    '''
    #Now we are going to collect all comments and replies of the discussion topics by looping through the dicussion topic ids
    for q in topic_list:
        q = str(q)
        url = "https://canvas.uva.nl/api/v1/"
        discus_url = url + "/courses" + "/" + class_id + "/discussion_topics" + "/" + q + "/entries"
        # Add token to header to authorize
        headers = {'Authorization': 'Bearer ' + str(token)}
        r = requests.get(discus_url, headers=headers, params={'per_page': 100})
        print ("got it", discus_url)
        r = r.json()
        print ("start writing data of "+ q + " dicussion board" + " to file")

        f = csv.writer(open('discussions.csv', 'a'))

        #Grabs certain data from the json to write to csv
        for i in r:
            print("found comment")
            f.writerow(['comment',
                        i['id'],
                        i['parent_id'],
                        i['created_at'],
                        i['updated_at'],
                        i['message'],
                        i['user_name'],
                        q])

            #As replies on messages are in a seperate place we grab them seperately
            if 'recent_replies' in i:
                    print("found replies")
                    for i in i['recent_replies']:
                        f.writerow(['replies',
                            i['id'],
                            i['parent_id'],
                            i['created_at'],
                            i['updated_at'],
                            i['message'],
                            i['user_name'],
                            q])

                    else:
                         print("No replies to process")
            print("Wrote data of discussion board " +q+" to file")
        print("finished writing all data. discussions.csv file ready to be opened")
    print("Finished all operation. I am tired.")
    print ("Going to sleep mode")
    print("Zzzzzzzz")

print(discussion_content("class_id","token"))
