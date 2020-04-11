// Create the main myMSALObj instance
// configuration parameters are located at authConfig.js
const myMSALObj = new Msal.UserAgentApplication(msalConfig);

function signIn() {
  myMSALObj.loginPopup(loginRequest)
    .then(loginResponse => {
      console.log('id_token acquired at: ' + new Date().toString());
      console.log(loginResponse);
      
      if (myMSALObj.getAccount()) {
        showWelcomeMessage(myMSALObj.getAccount());
      }
    }).catch(error => {
      console.log(error);
    });
}

function signOut() {
  myMSALObj.logout();
}

function getTokenPopup(request) {
  return myMSALObj.acquireTokenSilent(request)
    .catch(error => {
      console.log(error);
      console.log("silent token acquisition fails. acquiring token using popup");
          
      // fallback to interaction when silent call fails
        return myMSALObj.acquireTokenPopup(request)
          .then(tokenResponse => {
            return tokenResponse;
          }).catch(error => {
            console.log(error);
          });
    });
}

function seeProfile() {
  if (myMSALObj.getAccount()) {
    getTokenPopup(loginRequest)
      .then(response => {
        callMSGraph(graphConfig.graphMeEndpoint, response.accessToken, updateUI);
        profileButton.classList.add('d-none');
        mailButton.classList.remove('d-none');
        createMeetingButton.classList.remove('d-none')
      }).catch(error => {
        console.log(error);
      });
  }
}

function readMail() {
  if (myMSALObj.getAccount()) {
    getTokenPopup(tokenRequest)
      .then(response => {
        callMSGraph(graphConfig.graphMailEndpoint, response.accessToken, updateUI);
      }).catch(error => {
        console.log(error);
      });
  }
}

function createMeeting() {
    if (myMSALObj.getAccount()) {
      getTokenPopup(tokenRequest)
        .then(response => {

          var meetingDateTime = new Date(startDate.value + ' ' + startTime.value)

          var duration = document.querySelector('input[name="duration"]:checked').value
          if(duration === ""){
            duration = "15";
          }

          var time = meetingDateTime.getTime();
          time = time + (parseInt(duration)*60*1000);
          var meetingEnd = new Date(time);

          var subjectVal = subject.value;
          if(subjectVal === ""){
            subjectVal = "New Teams Meeting";
          }

          var createMeetingRequest =
          {
            "startDateTime" : meetingDateTime.toISOString(),
            "endDateTime" : meetingEnd.toISOString(),
            "subject" : subjectVal
          }         
    
          callMSGraphPost(graphConfig.graphMeetingEndpoint, response.accessToken, updateUI,createMeetingRequest);
        }).catch(error => {
          console.log(error);
        });
    }
}

