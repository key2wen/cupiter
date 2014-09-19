package com.to8to.weixin.util;

import com.to8to.commons.json.JSONObject;
import com.to8to.weixin.thrift.TSex;
import com.to8to.weixin.thrift.TUser;

public class WeixinUtil {
	public static void main(String[] args) {

	}

	// {"city":"上饶","country":"中国","headimgurl":"http:\/\/wx.qlogo.cn\/mmopen\/BnZvED9Hr1XB5pnpHYZ0G8PeRR3ycDf8SXKz6vAOXvoYjZmO4Cj8EXxc5fcpKCooEsDhibiaf9AGvAV2NRW9l3X3AlClT14cicG\/0",
	// "language":"zh_CN","nickname":"张云翔","openid":"o8n05t_I5FxOe6spiNbhB_ZI_zsg",
	// "province":"江西","remark":"","sex":1,"subscribe":1,"subscribe_time":1410527326}
	public static TUser getUserByJson(JSONObject json) {
		TUser user = new TUser();
		user.setOpenid(json.getString("openid"));
		user.setNickname(json.getString("nickname"));
		user.setSex(json.getInt("sex") == 1 ? TSex.MALE : TSex.FEMALE);
		user.setFollow_time(json.getLong("subscribe_time"));
		user.setProvince(json.getString("province"));
		user.setCity(json.getString("city"));
		user.setRemark(json.getString("remark"));
		user.setPic_header_url(json.getString("headimgurl"));
		return user;
	}

	/**
	 * 获取当时时间戳（10位）
	 */
	public static long getTimeStamp() {
		return System.currentTimeMillis() / 1000;
	}
}
