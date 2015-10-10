package com.xlstest;

import java.io.File;
import java.util.ArrayList;

import android.os.Bundle;
import android.view.View;
import android.view.View.OnClickListener;
import android.widget.Button;
import android.widget.TextView;
import android.app.Activity;
import android.content.SharedPreferences;

public class MainActivity extends Activity implements OnClickListener {

	private SharedPreferences spf;
	private boolean isFirst = true;
	private String path = "mnt/sdcard/XLSTest";

	private Button mk;
	private Button export;
	private Button read;
	private TextView show;

	private String content;

	@Override
	protected void onCreate(Bundle savedInstanceState) {
		super.onCreate(savedInstanceState);
		setContentView(R.layout.activity_main);
		findById();
		spf = getSharedPreferences("isFirst", 0);
		isFirst = spf.getBoolean("isFirst", true);
		if (isFirst == true) {
			setEnablefalse();
			spf.edit().putBoolean("isFirst", false).commit();
		} else {
			setEnableTrue();
		}
		setListener();
	}

	private void findById() {
		mk = (Button) findViewById(R.id.mk);
		export = (Button) findViewById(R.id.export);
		read = (Button) findViewById(R.id.read);
		show = (TextView) findViewById(R.id.show);
	}

	private void setListener() {
		mk.setOnClickListener(this);
		export.setOnClickListener(this);
		read.setOnClickListener(this);
	}

	private void setEnablefalse() {
		export.setEnabled(false);
		read.setEnabled(false);
	}

	private void setEnableTrue() {
		export.setEnabled(true);
		read.setEnabled(true);
	}

	private void mkDir() {
		File dir = new File(path);
		if (!dir.exists()) {
			dir.mkdir();
		}
	}

	private ArrayList<TestModel> getData() {
		ArrayList<TestModel> dataList = new ArrayList<TestModel>();
		for (int i = 1; i < 16; i++) {
			TestModel model = new TestModel();
			model.setId(String.valueOf(i));
			model.setDate("date:" + i);
			model.setTarget_temp_base("target_temp_base:" + i);
			model.setTarget_temp("target_temp:" + i);
			model.setEnvironmental_temp_base("environmental_temp_base:" + i);
			model.setEnvironmental_temp("environmental_temp:" + i);
			dataList.add(model);
		}
		return dataList;
	}

	@Override
	public void onClick(View v) {
		if (v == mk) {
			mkDir();
			export.setEnabled(true);
		} else if (v == export) {
			FileUtil.exportXLS(getData(), path + "/test");
			read.setEnabled(true);
		} else if (v == read) {
			content = FileUtil.readXLS(path + "/test.xls");
			show.setText(content);
		}
	}
}
