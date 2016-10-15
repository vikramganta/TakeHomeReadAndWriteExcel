/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.mycompany.ganta_excelreadwrite;

import java.util.Date;

/**
 *
 * @author s525718
 */
public class Song {
    private int sNo;
    private String genre;
    private int criticScore;
    private String albumName;
    private String artist;
    private Date releaseDate;

    public Song() {
    }

    public Song(int sNo, String genre, int criticScore, String albumName, String artist, Date releaseDate) {
        this.sNo = sNo;
        this.genre = genre;
        this.criticScore = criticScore;
        this.albumName = albumName;
        this.artist = artist;
        this.releaseDate = releaseDate;
    }

    public int getsNo() {
        return sNo;
    }

    public void setsNo(int sNo) {
        this.sNo = sNo;
    }

    public String getGenre() {
        return genre;
    }

    public void setGenre(String genre) {
        this.genre = genre;
    }

    public int getCriticScore() {
        return criticScore;
    }

    public void setCriticScore(int criticScore) {
        this.criticScore = criticScore;
    }

    public String getAlbumName() {
        return albumName;
    }

    public void setAlbumName(String albumName) {
        this.albumName = albumName;
    }

    public String getArtist() {
        return artist;
    }

    public void setArtist(String artist) {
        this.artist = artist;
    }

    public Date getReleaseDate() {
        return releaseDate;
    }

    public void setReleaseDate(Date releaseDate) {
        this.releaseDate = releaseDate;
    }
    
}
