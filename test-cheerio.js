const cheerio = require('cheerio');
const fs = require('fs');

const testHtml = `
<html>
<body>
<div id="corePriceDisplay_desktop_feature_div">
  <span class="a-price a-text-price"><span class="a-offscreen">£3.86</span></span>
  <span class="priceToPay"><span class="a-offscreen">£3.51</span></span>
</div>
<div id="merchant-info">
  Dispatched from Amazon<br>
  Sold by <a href="...">Bella's Hair and Beauty</a>
</div>
</body>
</html>
`;

const $ = cheerio.load(testHtml);

let amazonPrice = $('.priceToPay .a-offscreen').first().text().trim() ||
                  $('#corePriceDisplay_desktop_feature_div .priceToPay .a-offscreen').first().text().trim() ||
                  $('#corePrice_desktop .priceToPay .a-offscreen').first().text().trim() ||
                  $('#buyNew_noncbb .a-price .a-offscreen').first().text().trim() ||
                  $('#buyNewSection .a-price .a-offscreen').first().text().trim() ||
                  $('#desktop_buybox .a-price .a-offscreen').first().text().trim() ||
                  $('#price_inside_buybox').text().trim() ||
                  $('#corePriceDisplay_desktop_feature_div .a-price .a-offscreen').first().text().trim() ||
                  $('#corePrice_desktop .a-price .a-offscreen').first().text().trim() ||
                  $('.apex-core-price-identifier .a-offscreen').first().text().trim() ||
                  $('.apexPriceToPay .a-offscreen').first().text().trim() ||
                  $('.a-price.priceToPay .a-offscreen').first().text().trim() ||
                  $('.a-price .a-offscreen').first().text().trim() || "";
amazonPrice = amazonPrice.replace(/\s+/g, ' ').trim();

let amazonBuyboxOwner = $('div[tabular-attribute-name="Sold by"] .tabular-buybox-text').first().text().trim() ||
                        $('div[tabular-attribute-name="Verkauf durch"] .tabular-buybox-text').first().text().trim() ||
                        $('div[offer-display-feature-name="desktop-merchant-info"] .offer-display-feature-text-message').first().text().trim() ||
                        $('#sellerProfileTriggerId').first().text().trim() ||
                        $('#merchant-info a').first().text().trim();

if (!amazonBuyboxOwner) {
  let mInfo = $('#merchant-info').first().text().toLowerCase();
  if (mInfo.includes('sold by amazon') || mInfo.includes('verkauf durch amazon') || mInfo.includes('dispatched from and sold by amazon')) {
    amazonBuyboxOwner = "Amazon";
  } else {
    amazonBuyboxOwner = $('.offer-display-feature-text-message').first().text().trim();
  }
}
amazonBuyboxOwner = amazonBuyboxOwner.replace(/Sold by\s*:?\s*/gi, '').replace(/Venduto da\s*:?\s*/gi, '').replace(/Verkauf durch\s*:?\s*/gi, '').trim();

console.log("Price:", amazonPrice);
console.log("Buybox:", amazonBuyboxOwner);
