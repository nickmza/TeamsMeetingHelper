// Select DOM elements to work with
const welcomeDiv = document.getElementById("welcomeMessage");
const signInButton = document.getElementById("signIn");
const signOutButton = document.getElementById('signOut');
const cardDiv = document.getElementById("card-div");
const mailButton = document.getElementById("readMail");
const createMeetingButton = document.getElementById("createMeeting")
const profileButton = document.getElementById("seeProfile");
const profileDiv = document.getElementById("profile-div");
const createMeetingDiv = document.getElementById("createMeetingDiv");

const startDate = document.getElementById("startDate");
const startTime = document.getElementById("startTime");
const subject = document.getElementById("subject");
const duration = document.getElementById("duration");

const results = document.getElementById("resultsDiv");

function showWelcomeMessage(account) {

    // Reconfiguring DOM elements
    cardDiv.classList.remove('d-none');
    welcomeDiv.innerHTML = `Welcome ${account.name}`;
    signInButton.classList.add('d-none');
    signOutButton.classList.remove('d-none');

    seeProfile()
    setDefaults() 
}

function setDefaults(){
  var now = new Date();
  startDate.value = now.toISOString().split('T')[0]
  startTime.value = now.getHours() + ':' + now.getMinutes();
}

function updateUI(data, endpoint) {
  console.log('Graph API responded at: ' + new Date().toString());

  if (endpoint === graphConfig.graphMeEndpoint) {
    const title = document.createElement('p');
    title.innerHTML = "<strong>Title: </strong>" + data.jobTitle;
    const email = document.createElement('p');
    email.innerHTML = "<strong>Mail: </strong>" + data.mail;
    const phone = document.createElement('p');
    phone.innerHTML = "<strong>Phone: </strong>" + data.businessPhones[0];
    const address = document.createElement('p');
    address.innerHTML = "<strong>Location: </strong>" + data.officeLocation;
    profileDiv.appendChild(title);
    profileDiv.appendChild(email);
    profileDiv.appendChild(phone);
    profileDiv.appendChild(address);

    createMeetingDiv.classList.remove('d-none');
    
  } else if (endpoint == graphConfig.graphMeetingEndpoint){

    /*
    const resultsCard = document.getElementById("resultsCard");
    const meetingLink = document.createElement('a');
    meetingLink.setAttribute("href", data.joinWebUrl);
    meetingLink.innerText = "Teams Meeting";
    resultsCard.appendChild(meetingLink);
    */

    
    var linkHtml = decodeURIComponent(data.joinInformation.content)
    .replace(/\+/g," ")
    .replace("data:text/html,","")
    .trim();
    const linkContent =document.createElement('div');
    linkContent.innerHTML = linkHtml;
    results.appendChild(linkContent);
    

    results.classList.remove('d-none');
    
  }
}