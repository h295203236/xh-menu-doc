package org.mooon.biz;

import lombok.Value;

import java.util.Date;
import java.util.LinkedList;
import java.util.List;

public class MealMenu {
    String category;
    String forDay;
    Date forDayOfDate;
    List<Node> nodes;
    int priority;

    public MealMenu(String forDay, Date forDayOfDate) {
        this(forDay, forDayOfDate, null, 0);
    }

    public MealMenu(String forDay, Date forDayOfDate, String category, int priority) {
        this.category = category;
        this.forDay = forDay;
        this.forDayOfDate = forDayOfDate;
        this.nodes = new LinkedList<>();
        this.priority = priority;
    }

    public Date getForDayOfDate() {
        return forDayOfDate;
    }

    public String getForDay() {
        return forDay;
    }

    public void setPriority(int priority) {
        this.priority = priority;
    }

    public void setCategory(String category) {
        this.category = category;
    }

    public void addMenu(String name, String price) {
        this.nodes.add(new Node(name, price));
    }

    public List<Node> menuList() {
        return nodes;
    }

    @Value
    public static class Node {
        String name;
        String price;
    }
}
