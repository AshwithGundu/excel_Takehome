/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package take_home_gundu;

import java.util.Comparator;

/**
 *
 * @author S
 */
public class data  {
    
    String album_Name, genre, artist, release_date;
    
    Integer  credit_score;

    public String getAlbum_Name() {
        return album_Name;
    }

    public void setAlbum_Name(String album_Name) {
        this.album_Name = album_Name;
    }

    public String getGenre() {
        return genre;
    }

    public void setGenre(String genre) {
        this.genre = genre;
    }

    public String getArtist() {
        return artist;
    }

    public void setArtist(String artist) {
        this.artist = artist;
    }

    public String getRelease_date() {
        return release_date;
    }

    public void setRelease_date(String release_date) {
        this.release_date = release_date;
    }

    public Integer getCredit_score() {
        return credit_score;
    }

    public void setCredit_score(Integer credit_score) {
        this.credit_score = credit_score;
    }

    
    
    public data(){}

    public data(String album_Name, String genre, String artist, String release_date, int credit_score) {
        super();
        this.album_Name = album_Name;
        this.genre = genre;
        this.artist = artist;
        this.release_date = release_date;
        this.credit_score = credit_score;
    }

    @Override
    public String toString() {
        return "{" + "album_Name=" + album_Name + ", genre=" + genre + ", artist=" + artist + ", release_date=" + release_date + ", credit_score=" + credit_score + '}';
    }
  

    
}
