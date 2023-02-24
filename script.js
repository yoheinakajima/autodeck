


async function onOpen(e) {
  var response = await createDeck("Kitten mittens");

}


function appendSlideWithTitleAndBody(title, body) {
  var presentation = SlidesApp.getActivePresentation();
  var slide = presentation.appendSlide();
  
  var slideWidth = presentation.getPageWidth();
  var slideHeight = presentation.getPageHeight();
  
  var titleShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, 50, 50, slideWidth - 100, 50);
  var titleText = titleShape.getText();
  titleText.setText(title);
  titleText.getTextStyle().setFontSize(36);
  titleText.getTextStyle().setBold(true);
  
  var bodyShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, 50, 120, slideWidth - 100, slideHeight - 150);
  var bodyText = bodyShape.getText();
  bodyText.setText(body);
  bodyText.getTextStyle().setFontSize(30);
}


function appendSlideWithTitleBodyAndImage(title, body, imageUrl) {
  var presentation = SlidesApp.getActivePresentation();
  var slide = presentation.appendSlide();
  
  var slideWidth = presentation.getPageWidth();
  var slideHeight = presentation.getPageHeight();
  
  var titleShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, 50, 50, slideWidth - 300, 50);
  titleShape.getText().setText(title);
  
  var bodyShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, 50, 100, slideWidth - 300, slideHeight - 150);
  bodyShape.getText().setText(body);
  
  var imageWidth = slideWidth - (slideWidth - 250);
  var imageHeight = slideHeight - 100;
  
  var imageBlob = UrlFetchApp.fetch(imageUrl).getBlob();
  var image = slide.insertImage(imageBlob, slideWidth - imageWidth - 50, 50, imageWidth, imageHeight);
}




async function createDeck(input) {
  updateFirstShapeText(input);
  console.log("getCompletion");
  const apiUrl = 'https://api.openai.com/v1/completions';
  prompt = 'You are an illustrative writer and expert marketer who specializes in venture capital and startup analysis. For the provided startup description, write a 7 part narrarative targeting investors on why investing in this startup is great opportunity. Each section should be one sentence. Provide the answer as a JSON array with the following format:{"problem":"This startup provides a platform to create an online marketplace for small independent sellers to connect with customers, reducing barriers to entry to the e-commerce space and allowing them to reach more customers than ever before.","approach":"By leveraging a network of small business owners, the startup is empowering these merchants to compete with the established players in the e-commerce space, leveling the playing field and providing better value to customers.","solution":"The startup\'s platform offers an easy-to-use interface and low transaction fees for merchants, allowing them to quickly launch their own online store and start selling in minutes.","customer":"The ideal customer for this startup is a small business owner who wants to reach more customers but doesn\u2019t have the resources to set up their own website or pay for expensive advertising.","market":"The total addressable market for this startup is estimated to be around $10 billion globally, with the potential to expand into other markets as the platform becomes more popular.","go to market":"The startup will focus on onboarding small business owners through targeted ads and influencer marketing, as well as building relationships with independent retailers and providing them with the resources they need to get started.","funding":"The startup is looking to raise $3 million in venture capital, with the goal of selling 10-15% of the company for this amount.","inspiring quote":"\'Today\u2019s mighty oak is just yesterday\u2019s nut, that held its ground\' - David Icke"}### STARTUP DESCRIPION'+input+'.###RESULT (JSON):';
  const body = JSON.stringify({
    model : "text-davinci-003",
    prompt: prompt,
    max_tokens: 1000,
    stop: ['###'],
    temperature: 0.7
  });
  const options = {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer sk-...` //ADD API KEY HERE
    },
    payload: body,
    muteHttpExceptions: true
  };
  const response = UrlFetchApp.fetch(apiUrl, options);
  const json = JSON.parse(response.getContentText());
  console.log(json);
  const answer = JSON.parse(json.choices[0].text.trim());
  processJson(answer);
  return json.choices[0].text.trim();
}


function processJson(json) {
  console.log("processJson")
  for (const key in json) {
    if (json.hasOwnProperty(key)) {
      const value = json[key];
      // Call your function here and pass in the key and value
      // For example, if your function is called myFunction:
      //myFunction(key, value);
      console.log("key: "+key+". value: "+value);
      appendSlideWithTitleAndBody(key, value);
    }
  }
}


function updateFirstShapeText(text) {
  var presentation = SlidesApp.getActivePresentation();
  var slide = presentation.getSlides()[0]; // Get the first slide
  var shape = slide.getShapes()[0]; // Get the first shape of the slide
  var textRange = shape.getText();
  textRange.setText(text); // Update the text of the shape
  var shape = slide.getShapes()[1]; // Get the first shape of the slide
  var textRange = shape.getText();
  textRange.setText("Pitch Deck Draft"); // Update the text of the shape
}
