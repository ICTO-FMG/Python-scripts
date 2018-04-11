import requests

## Read token from file located in .secret/token
with open('secret/token', 'r') as f: 
	token = f.readline()

# Add token to header to authorize
headers = {'Authorization': 'Bearer ' + str(token)}

def sections_to_groups(name_groupset, courseID, headers):
    """
    Creates a Group in Canvas for each Section in a course and adds all students enrolled in a Section to the corresponding Group.
    Students can not be enrolled in more than one Section. A Section can not have more than 100 students enrolled (because of pagination).
    
    name_groupset (string): name of the groupset you want to create
    courseID (integer): courseID of the course (see Canvas URL)
    headers (dict): authorisation header
    """
    
    url = "https://canvas.uva.nl/api/v1/"

    # create group category
    url_groupcat = url + 'courses/' + str(courseID) + '/group_categories'
    groupcat = {'name':name_groupset,'self_signup':'null'}
    r = requests.post(url_groupcat, headers=headers, params=groupcat)
    created_groupcat = r.json()

    if 'errors' in created_groupcat:
        print "Error while creating group category:"
        print created_groupcat['errors']
        return False

    groupcatID = created_groupcat['id']
    print "group cat ID", groupcatID
    
    # get sections
    url_sections = url + "courses/" + str(courseID) + "/sections"
    r =  requests.get(url_sections, headers=headers, params={'per_page': 100})
    sections = r.json()

    for section in sections:
        section_name = section['name']
        sectionID = section['id']
        print "Section ID:", sectionID

        # create group for this section
        url_groups = url + "group_categories/" + str(groupcatID) + "/groups"

        group = {'name':section_name, 'join_level':'parent_context_auto_join'}
        r = requests.post(url_groups, headers=headers, params=group)
        
        groupID = r.json()['id']
        print "group ID:", groupID, section_name
        
        # get users from the section
        url_section_users = url + "sections/" + str(sectionID) + "/enrollments"
        r =  requests.get(url_section_users, headers=headers, params={'per_page': 100})
        users = r.json()
        
        userids = []
        for user in users:
            if user['role'] == "StudentEnrollment":
                userids.append(str(user['user_id']))
                
        if not users:
            print "no users in Section", section_name
            continue

        # enroll users to group
        url_membership = url + "groups/" + str(groupID) + "/memberships"
        
        for userID in userids:
            user = {"user_id": str(userID)}
            print "user ID:", userID
            r = requests.post(url_membership, headers=headers, params=user)
            
        print len(userids), "students added to group", section_name
        
    print "Done"           
    return True

# Example:
#sections_to_groups("Seminar groups", 11111111, headers)
