require 'builder'

class UploadController < ApplicationController
  def create
    file = params[:file]
    filename = params[:filename].presence || File.basename(file.original_filename, ".*")

    Spreadsheet.client_encoding = 'UTF-8'
    spreadsheet = Spreadsheet.open file.open
    xml = generate_xml(spreadsheet)
    raise "DebitAmt and CreditAmt does not tally." if @total_debit.round(2) != @total_credit.round(2)
    raise "HomeDebitAmt and HomeCreditAmt does not tally." if @total_home_debit.round(2) != @total_home_credit.round(2)

    send_data(xml, type: "text/xml", filename: filename)
  end

  private
  def generate_xml(spreadsheet)
    @total_debit = 0
    @total_credit = 0
    @total_home_debit = 0
    @total_home_credit = 0
    builder = Builder::XmlMarkup.new(:target => "", :indent => 1)
    builder.target!
    sheet1 = spreadsheet.worksheet 0
    builder.GLInterface do
      sheet1.each_with_index do |row, index|
        if index != 0
          @total_debit += row[8].to_f
          @total_credit += row[9].to_f
          @total_home_debit += row[11].to_f
          @total_home_credit += row[12].to_f

          if all_columns_filled?(row)
            builder.GLInterfaceRow do |b|
              ENV['column_names'].each_with_index do |column, index|
                b.send("#{column}=".to_sym, row[index])
              end
              b.GLEntity(row[0])
              b.GLCode(row[1])
              b.EntityCode(row[2])
              b.Year(row[3])
              b.Period(row[4])
              b.Reference(row[5])
              b.Explanation(row[6])
              b.TransDate(row[7])
              b.DebitAmt(row[8])
              b.CreditAmt(row[9])
              b.CurrencyCode(row[10])
              b.HomeDebitAmt(row[11])
              b.HomeCreditAmt(row[12])
            end
          end
        end
      end
    end
  end

  def all_columns_filled?(row)
    (0..12).to_a.each do |i|
      return false if row[i].to_s.length <= 0
    end
  end
end