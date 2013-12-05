/*
******************************************************
* File:Skype.js
* Date:26/8/2013
* Version: 1.1
* Author: Alban Vidal-Naquet (alban@albanvn.net)
* Sarah plugin for skype
******************************************************
*/

var bf=require("./basicfunctions.js");

// Constantes
const cst_wscript="wscript.exe";
const cst_skypevbs="skype.vbs";
const cst_msg_unknowlast="Je ne sais pas de qui vous parler";
const cst_msg_calling="Très bien, j'appelle";
const cst_msg_callingvideo1="Très bien, j'appelle";
const cst_msg_callingvideo2="et lance la vidéo";
const cst_msg_whoisconnected="Je regarde qui est connecté";
const cst_msg_finishcall="Très bien je raccroche";
const cst_msg_isconnected_b="Je regarde si";
const cst_msg_isconnected_e="est connécté";
const cst_msg_getstatus="Je regarde votre statut skaillpe";
const cst_msg_connect="Je connecte votre compte skaillpe";
const cst_msg_okletstry="Je me renseigne";
const cst_msg_busy="Je mets en occupé votre compte skaillpe";
const cst_msg_disconnect="Je déconnecte votre compte skaillpe";
const cst_msg_selectfriend="Je mets à jour votre liste d'amis";
const cst_msg_connectaccount="Je connecte le compte skaillpe";
const cst_msg_badconfiguration="Veuillez configurer correctement les paramètres du pleugue ine";
const cst_msg_fullscreen="Très bien je mets skaillpe en plein écran";
const cst_msg_unknownaccount="Je ne connais pas le compte spécifié";
const cst_msg_missedcall="Appel reçu de";
const cst_msg_missedcall2="le";
const cst_msg_nomissedcall="Aucun appel manqué";
const cst_msg_missedcallintro="Voici la liste des appels manqués.";
const cst_msg_missedcallintro2="Oui on a essayé de vous joindre";
const cst_maxtimeout=20*1000;
const cst_minperiod_sameid=5*60*1000;
const cst_month=["janvier","février","mars","avril","mai","juin","juillet","aout","septembre","octobre","novembre","décembre"];

// Global variable
var g_script_skype_path="";
var	g_timeout=0;
var g_lastisconnected="";
var g_status="";
var g_account="";
var g_call=new Array();
var g_missedcall=new Array();

exports.init = function (SARAH)
{
	var cfg=SARAH.ConfigManager.getConfig();
	cfg = cfg.modules.skype;
	g_script_skype_path=__dirname + "\\" + cst_skypevbs;
	SARAH.remote({'run' : cst_wscript, 'runp' : g_script_skype_path });
	var parameter="updateparameter \"" + __dirname + "\" \"" +cfg.Name1+"\" \""+cfg.Name2+"\" \""+cfg.Name3+"\"";
	SARAH.remote({'run' : cst_wscript, 'runp' : g_script_skype_path + " " + parameter });
}

exports.release = function (SARAH)
{
	SARAH.remote({'run' : cst_wscript, 'runp' : g_script_skype_path + " " + "cleanwscript"});
}

exports.action = function(data, callback, config, SARAH)
{
	if (data.mode=="call" || data.mode=="callvideo" || data.mode=="isconnected" || data.mode=="calllast" || data.mode=="calllastvideo")
	{
	   if (data.mode=="calllast" || data.mode=="calllastvideo")
	   {
		 if (new Date().getTime()<g_timeout)
		 {
		  data.name=g_lastisconnected;
		  data.mode="call";
		 }
		 else
		 {
		   data.mode="";
		   bf.speak(cst_msg_unknowlast, SARAH);
		 }
	   }		 
	   if (data.mode=="callvideo)") bf.speak(cst_msg_callingvideo1 + " " + data.name + " " + cst_msg_callingvideo2, SARAH);
	   if (data.mode=="call") bf.speak(cst_msg_calling + " " + data.name, SARAH);
	   if (data.mode=="isconnected")
	   {
		 bf.speak(cst_msg_isconnected_b + " " + data.name + " " + cst_msg_isconnected_e, SARAH);
		 g_timeout=new Date().getTime()+cst_maxtimeout;
		 g_lastisconnected=data.name;
	   }
	   else
		 timer=0;
	  if (data.mode!="")
		SARAH.remote({ 'run' : cst_wscript, 'runp' : g_script_skype_path + " " + data.mode + " \"" + data.name + "\""});
	}
	else if (data.mode=="connectaccount")
	{
	  var cfg=SARAH.ConfigManager.getConfig();
	  cfg = cfg.modules.skype;
	  if (cfg.Skype_path!="") 
	  {
		exe = cfg.Skype_path + "\\" + "Skype.exe";
		SARAH.remote({ 'run' : exe, 'runp' : "/shutdown"});
		var login="";
		var password="";
		var name="";
		switch (data.account)
		{
			case "1":
				login=cfg.User1;
				password=cfg.Pass1;
				name=cfg.Name1;
				break;
			case "2":
				login=cfg.User2;
				password=cfg.Pass2;
				name=cfg.Name2;
				break;
			case "3":
				login=cfg.User3;
				password=cfg.Pass3;
				name=cfg.Name3;
				break;
			default:
				console.log("Unknown account #"+data.account);
				break;
		}
		if (login!="" && password!="" && name!="")
		{
			bf.speak(cst_msg_connectaccount + " " + "\"" + name + "\"", SARAH);
			setTimeout(function(){SARAH.remote({ 'run' : exe, 'runp' : "\"/username:" + login + "\" \"" + "/password:" + password + "\""});
							  setTimeout(function(){SARAH.remote({ 'run' : cst_wscript, 'runp' : g_script_skype_path + " selectfriendsilent " + __dirname + " " + "\"" + cfg.Skype_list + "\""});
												   },10000);
							 },2000);
		}
		else
		  bf.speak(cst_msg_unknownaccount, SARAH);
	 }
	  else
		bf.speak(cst_msg_badconfiguration, SARAH);
	}
	else if (data.mode=="status")
	{
	   if (typeof data.account!='undefined' && data.account!="") g_account=data.account;
	   if (typeof data.status!='undefined' && data.status!="") g_status=data.status;
	   if (typeof data.lastcall!='undefined' && data.lastcall!="")
	   {
		 var skip=0;

		 SARAH.play(__dirname+"\\sonnerie.mp3");
		 // If it's the same contact than the last one on short period then ignore it
		 if (g_call.length>0 && g_call[g_call.length-1].id==data.lastcall && new Date.getTime()<(g_call[g_call.length-1].date.getTime()+cst_minperiod_sameid))
			 skip=1;
		 if (skip==0)
		 {
			ref=new Date();
			g_call.push({id:data.lastcall, date:ref});
			// Current call may be missed, flag it as missed for the moment
			g_missedcall.push({id:data.lastcall, date:ref});
		 }
	   }
	}
	else if (data.mode=="lastmissedcalls")
	{
	  if (g_missedcall.length>0)
	  {
		bf.speak(cst_msg_missedcallintro, SARAH);
		for (i=0;i<g_missedcall.length;i++)
		  bf.speak(cst_msg_missedcall + " " + g_missedcall[i].id + " " + cst_msg_missedcall2 + " " + formatDate(g_missedcall[g_call.length-1].date,1), SARAH);
	  }
	  else
		bf.speak(cst_msg_nomissedcall, SARAH);
	}
	else if (data.mode=="test")
	{
	  bf.speak(cst_msg_okletstry, SARAH);
	  SARAH.remote({ 'run' : cst_wscript, 'runp' : g_script_skype_path + " " + data.mode + " " + __dirname});
	}
	else
	{
	  var cfg=SARAH.ConfigManager.getConfig();
	  var text="";
	  cfg = cfg.modules.skype;
	  optionnal="";
	  if (data.mode=="selectfriend")
	  { 
		text=cst_msg_selectfriend;
		optionnal=" \"" + cfg.Skype_list + "\"";
	  }
	  if (data.mode=="getstatus") text=cst_msg_getstatus;
	  if (data.mode=="connect") text=cst_msg_connect;
	  if (data.mode=="busy") text=cst_msg_busy;
	  if (data.mode=="disconnect") text=cst_msg_disconnect;
	  if (data.mode=="listconnected") text=cst_msg_whoisconnected;
	  if (data.mode=="finish") text=cst_msg_finishcall;
	  if (data.mode=="fullscreen") text=cst_msg_fullscreen;
	  if (data.mode=="answer")
		 // forget about the current call, this is not a missed one...	
		 g_missedcall.pop();
	  if (data.mode!="") 
	  {
		  if (text!="") 
			bf.speak(text, SARAH);
		  SARAH.remote({ 'run' : cst_wscript, 'runp' : g_script_skype_path + " " + data.mode + " " + __dirname + optionnal});
	  }
	}
	callback();
}


var formatDate=function(d, tovocalize)
{
  str="";
  if (tovocalize==0)
	str=d.getDate()+"/"+(d.getMonth()+1)+"/"+d.getFullYear()+" "+d.getHours()+":"+d.getMinutes();
  else
	str=d.getDate()+" "+cst_month[d.getMonth()]+" "+d.getFullYear()+" à "+d.getHours()+" heures "+d.getMinutes()+" minutes";
  return str;
}

exports.getBasic = function(SARAH)
{
  info={};
  info.account=g_account;
  info.status=g_status;
  info.lastmissedcall={};
  if (g_call.length>0)
	info.lastcall=g_call[g_call.length-1];
  else
	info.lastcall={id:"",date:""};
  if (g_missedcall.length>0)
  {
	info.lastmissedcall.id=g_missedcall[g_call.length-1].id;
	info.lastmissedcall.date=formatDate(g_missedcall[g_call.length-1].date,0);
  }
  else
	info.lastmissedcall={id:"",date:""};
  return info;
}
