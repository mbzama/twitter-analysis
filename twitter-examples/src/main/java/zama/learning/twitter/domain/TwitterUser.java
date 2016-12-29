package zama.learning.twitter.domain;

import zama.learning.twitter.excel.annotation.ExcelGrid;

public class TwitterUser {
	private String id;
	private String name;
	private String screenName;
	private String location;
	private String description;

	public TwitterUser(String id, String name, String screenName, String location, String description){
		this.id = id;
		this.name = name;
		this.screenName = screenName;
		this.location = location;
		this.description = description;
	}

	@ExcelGrid(header="Id", order=1)
	public String getId() {
		return id;
	}

	public void setId(String id) {
		this.id = id;
	}

	@ExcelGrid(header="Name", order=2)
	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	@ExcelGrid(header="Screen Name", order=3)
	public String getScreenName() {
		return screenName;
	}

	public void setScreenName(String screenName) {
		this.screenName = screenName;
	}

	@ExcelGrid(header="Location", order=4)
	public String getLocation() {
		return location;
	}

	public void setLocation(String location) {
		this.location = location;
	}

	@ExcelGrid(header="Description", order=5)
	public String getDescription() {
		return description;
	}

	public void setDescription(String description) {
		this.description = description;
	}

	@Override
	public String toString() {
		return "TwitterUser [id=" + id + ", name=" + name + ", screenName=" + screenName + ", location=" + location
				+ ", description=" + description + "]";
	}
}	
