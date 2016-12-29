package zama.learning.twitter;

import java.io.File;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import twitter4j.Location;
import twitter4j.PagableResponseList;
import twitter4j.RateLimitStatus;
import twitter4j.ResponseList;
import twitter4j.Twitter;
import twitter4j.TwitterException;
import twitter4j.TwitterFactory;
import twitter4j.User;
import twitter4j.auth.OAuth2Token;
import twitter4j.conf.ConfigurationBuilder;
import zama.learning.twitter.constants.TwitterContants;
import zama.learning.twitter.domain.TwitterUser;
import zama.learning.twitter.excel.ExcelGenerator;

public class TwitterDemo {

	public static void main(String[] args) throws TwitterException {
		Twitter twitter = getTwitter();
		Map<String, RateLimitStatus> rateLimitStatus = twitter.getRateLimitStatus("search");
		RateLimitStatus searchTweetsRateLimit = rateLimitStatus.get("/search/tweets");
		System.out.printf("You have %d calls remaining out of %d, Limit resets in %d seconds\n",
				searchTweetsRateLimit.getRemaining(), searchTweetsRateLimit.getLimit(),
				searchTweetsRateLimit.getSecondsUntilReset());
		
		List<TwitterUser> twitterUsers = getTwitterFollowers(twitter);
		writeResultsToExcelFile(twitterUsers);
		
		//getAvailableTrends(twitter);
	}


	private static void writeResultsToExcelFile(List<TwitterUser> twitterUsers) {
		ExcelGenerator<TwitterUser> excelGenerator = new ExcelGenerator<>();
		HSSFWorkbook workbook = excelGenerator.generateExcel(twitterUsers);
		String filePath = "C:\\Users\\bmohammad\\Desktop\\temp\\";
		String fileName = "Twitter_Followers_"+Calendar.getInstance().getTimeInMillis()+".xls";
		File file = new File(filePath+fileName);
		excelGenerator.writeFile(workbook, file);
	}


	private static List<TwitterUser> getTwitterFollowers(Twitter twitter) throws TwitterException {
		List<TwitterUser> twitterUsers = new ArrayList<>();
		PagableResponseList<User> followers = null;
		long nextCursor = -1;
		
		try {
			followers = twitter.getFollowersList("EvokeUS", nextCursor, 200);
			System.out.println("Prev Cursor: "+followers.getPreviousCursor());
			System.out.println("Next Cursor: "+followers.getNextCursor());
			System.out.println("Page Items: "+followers.size()+"---------\n");
			for (User user : followers) {
				System.out.println(user);
				twitterUsers.add(new TwitterUser(String.valueOf(user.getId()), user.getName(), user.getScreenName(), user.getLocation(), user.getDescription()));
			}
			
			nextCursor = followers.getNextCursor();
			System.out.println("nextCursor: "+nextCursor);
			while(nextCursor != 0){
				System.out.println("While loop..");
				followers = twitter.getFollowersList("EvokeUS", nextCursor, 200);
				for (User user : followers) {
					System.out.println(user);
					twitterUsers.add(new TwitterUser(String.valueOf(user.getId()), user.getName(), user.getScreenName(), user.getLocation(), user.getDescription()));
				}
				
				System.out.println("Prev Cursor: "+followers.getPreviousCursor());
				System.out.println("Next Cursor: "+followers.getNextCursor());
				nextCursor = followers.getNextCursor();
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		
		System.out.println("Twitter Followers: "+twitterUsers.size());
		return twitterUsers;
	}


	private static void getAvailableTrends(Twitter twitter) {
		try {
			ResponseList<Location> locations;
			locations = twitter.getAvailableTrends();
			System.out.println("Showing available trends");
			for (Location location : locations) {
				System.out.println(location.getName() + " (woeid:" + location.getWoeid() + ")");
			}
		} catch (TwitterException te) {
			te.printStackTrace();
			System.out.println("Failed to get trends: " + te.getMessage());
			System.exit(-1);
		}
	}

	private static OAuth2Token getOAuth2Token() {
		OAuth2Token token = null;
		ConfigurationBuilder cb;

		cb = new ConfigurationBuilder();
		cb.setApplicationOnlyAuthEnabled(true);

		cb.setOAuthConsumerKey(TwitterContants.ZAMA_CONSUMER_KEY).setOAuthConsumerSecret(TwitterContants.ZAMA_CONSUMER_KEY_SECRET);

		try {
			token = new TwitterFactory(cb.build()).getInstance().getOAuth2Token();
		} catch (Exception e) {
			e.printStackTrace();
			System.exit(0);
		}

		return token;
	}

	private static Twitter getTwitter() {
		OAuth2Token token;

		// First step, get a "bearer" token that can be used for our requests
		token = getOAuth2Token();

		ConfigurationBuilder cb = new ConfigurationBuilder();

		cb.setApplicationOnlyAuthEnabled(true);

		cb.setOAuthConsumerKey(TwitterContants.ZAMA_CONSUMER_KEY);
		cb.setOAuthConsumerSecret(TwitterContants.ZAMA_CONSUMER_KEY_SECRET);

		cb.setOAuth2TokenType(token.getTokenType());
		cb.setOAuth2AccessToken(token.getAccessToken());

		// And create the Twitter object!
		return new TwitterFactory(cb.build()).getInstance();
	}
}
