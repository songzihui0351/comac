package main;

import lombok.Data;

import java.util.Arrays;
import java.util.HashSet;
import java.util.Objects;

@Data
public class Attendance implements Comparable<Attendance> {
    private String name;
    private String dept;
    private String membership;
    private String city;
    private String hotel;
    private String[] dateList;
    private HashSet<String> dateSet;
    private String arrivalDate;
    private String departureDate;

    Attendance(String name, String dept, String membership, String date, int days) {
        this.name = name;
        this.dept = dept;
        this.membership = membership;
        this.dateSet = new HashSet<>();
        dateList = new String[days];
        if (date != null) {
            dateSet.add(date);
        }
    }

    public Attendance(String name, String dept, String hotel) {
        this.name = name;
        this.dept = dept;
        this.hotel = hotel;
    }

    public Attendance(String name, String dept, String hotel, String city, String membership, String date, Integer days) {
        this.name = name;
        this.dept = dept;
        this.city = city;
        this.hotel = hotel;
        this.membership = membership;
        this.dateSet = new HashSet<>();
        if (date != null) {
            dateSet.add(date);
        }
        dateList = new String[days];
        Arrays.fill(dateList, "其");
        int i = ((Integer.parseInt((date.split("\\.")[1]))) - 26 + days) % days;
        dateList[i] = city;
    }

    @Override
    public boolean equals(Object obj) {
        Attendance that = (Attendance) obj;
        return Objects.equals(name, that.name) &&
                Objects.equals(dept, that.dept);
    }

    @Override
    public int hashCode() {
        return Objects.hash(name, dept);
    }

    @Override
    public int compareTo(Attendance o) {
        return this.name.compareTo(o.name);
    }
}
