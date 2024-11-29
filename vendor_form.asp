<div class="form-group">
    <label for="ContactPerson">客服聯絡窗口：</label>
    <input type="text" id="ContactPerson" name="ContactPerson" 
           value="<%=rs("ContactPerson")%>" required>
</div>

<div class="form-group">
    <label for="LogisticsContact">物流聯絡窗口：</label>
    <input type="text" id="LogisticsContact" name="LogisticsContact" 
           value="<%=rs("LogisticsContact")%>">
</div>

<div class="form-group">
    <label for="MarketingContact">行銷聯絡窗口：</label>
    <input type="text" id="MarketingContact" name="MarketingContact" 
           value="<%=rs("MarketingContact")%>">
</div> 