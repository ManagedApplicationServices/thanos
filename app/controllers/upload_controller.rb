require 'builder'

class UploadController < ApplicationController
  def create
    file = params[:file]
    filename = params[:filename].presence || File.basename(file.original_filename, ".*")

    Spreadsheet.client_encoding = 'UTF-8'
    spreadsheet = Spreadsheet.open file.open
    xml = generate_xml(spreadsheet)
    raise "Row #{@invalid_rows.join(", ")} is a invalid row." if @invalid_rows.size > 0
    raise "DebitAmt and CreditAmt does not tally." if @total_debit.round(2) != @total_credit.round(2)
    raise "HomeDebitAmt and HomeCreditAmt does not tally." if @total_home_debit.round(2) != @total_home_credit.round(2)

    send_data(xml, type: "text/xml", filename: filename)
  end

  private
  def generate_xml(spreadsheet)
    @invalid_rows = []
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
          if all_columns_filled?(row)
            @total_debit += row[8].to_f
            @total_credit += row[9].to_f
            @total_home_debit += row[11].to_f
            @total_home_credit += row[12].to_f

            builder.GLInterfaceRow do |b|
              ENV["column_names"].split(",").each_with_index do |column, index|
                b.__send__(column, row[index])
              end
            end
          else
            @invalid_rows << index + 1
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