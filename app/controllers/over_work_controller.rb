require 'docx'
require 'tempfile'

class OverWorkController < ApplicationController
  def generate_overtime_report

  end

  def users_by_date
    start_date = Date.parse(params[:start_date]) rescue nil
    end_date = Date.parse(params[:end_date]) rescue nil
  
    if start_date.nil? || end_date.nil?
      render json: [], status: :unprocessable_entity and return
    end
  
    custom_field = TimeEntryCustomField.find_by(name: 'Тип работ')
    overtime_entries = TimeEntry
      .joins(:custom_values)
      .where(spent_on: start_date..end_date)
      .where(custom_values: {
        custom_field_id: custom_field.id,
        value: 'Сверхурочная'
      })
  
    users_data = {}
  
    overtime_entries.each do |entry|
      user = entry.user
      users_data[user.id] ||= {
        id: user.id,
        name: "#{user.firstname} #{user.lastname}",
        entries: {}
      }
      date_str = entry.spent_on.to_s
      users_data[user.id][:entries][date_str] ||= 0.0
      users_data[user.id][:entries][date_str] += entry.hours
    end
  
    result = users_data.values.map do |user|
      {
        id: user[:id],
        name: user[:name],
        entries: user[:entries].map { |date, hours| { date: date, hours: hours } }
      }
    end
  
    render json: result
  end
  
  
  

  def do_generate_ov
    user_ids = params[:user_ids] || []
    @global_user_ids = user_ids
    start_date_str = params[:start_date]
    end_date_str = params[:end_date]
    type_overwork = params[:option_select]
    report_type = params[:report_type]

    @global_type_overwork = type_overwork.to_s

    if user_ids.blank?
      flash[:error] = "Пользователи не выбраны"
      redirect_to action: :generate_overtime_report and return 
    end

    if start_date_str.blank?
      flash[:error] = "Укажите дату начала"
      redirect_to action: :generate_overtime_report and return 
    end

    if end_date_str.blank?
      flash[:error] = "Укажите дату окончания"
      redirect_to action: :generate_overtime_report and return 
    end 

    start_date = Date.parse(start_date_str)
    end_date = Date.parse(end_date_str)

    date_range = (start_date..end_date).to_a
    @date_range_overtime = date_range

    custom_field = TimeEntryCustomField.find_by(name: 'Тип работ')

    entries = TimeEntry
      .includes(:issue)
      .joins(:custom_values)
      .where(user_id: user_ids)
      .where(spent_on: date_range)
      .where(custom_values: {
        custom_field_id: custom_field.id,
        value: 'Сверхурочная'
      })

    @overtime_issues = entries.map(&:issue).compact.uniq
    @users = User.where(id: user_ids)

    if @overtime_issues.blank?
      flash[:error] = "У выбранных пользователей нет задач с типом работ 'Сверхурочная' за выбранный период."
      redirect_to action: :generate_overtime_report and return
    end

    respond_to do |format|


      if report_type == "dokladnaya"
        format.xlsx do
          response.headers['Content-Disposition'] = 'attachment; filename=Докладная.xlsx'
          render xlsx: 'Докладная', template: 'over_work/do_dokladnaya'
        end
      else
        format.xlsx do
          response.headers['Content-Disposition'] = 'attachment; filename=Отчет.xlsx'
          render xlsx: 'Отчет о сверхурочной работе', template: 'over_work/do_generate_ov'
        end
      end
    end
  end


  # def do_dokladnaya
  # user_ids = params[:user_ids] || []
  #   @global_user_ids = user_ids
  #   start_date_str = params[:start_date]
  #   end_date_str = params[:end_date]
  #   type_overwork_dokladnaya = params[:option_select]

  #   @global_type_overwork_dokladnaya = type_overwork_dokladnaya.to_s
    

  #   if user_ids.blank?
  #     flash[:error] = "Пользователи не выбраны"
  #     redirect_to action: :generate_overtime_report and return 
  #   end

  #   if start_date_str.blank?
  #     flash[:error] = "Укажите дату начала"
  #     redirect_to action: :generate_overtime_report and return 
  #   end

  #   if end_date_str.blank?
  #     flash[:error] = "Укажите дату окончания"
  #     redirect_to action: :generate_overtime_report and return 
  #   end 

  #   start_date = Date.parse(start_date_str)
  #   end_date = Date.parse(end_date_str)

  #   date_range = (start_date..end_date).to_a
  #   @date_range_overtime = date_range

  #   custom_field = TimeEntryCustomField.find_by(name: 'Тип работ')

  #   entries = TimeEntry
  #     .includes(:issue)
  #     .joins(:custom_values)
  #     .where(user_id: user_ids)
  #     .where(spent_on: date_range)
  #     .where(custom_values: {
  #       custom_field_id: custom_field.id,
  #       value: 'Сверхурочная'
  #     })

  #   @overtime_issues = entries.map(&:issue).compact.uniq
  #   @users = User.where(id: user_ids)

  #   respond_to do |format|
  #     format.xlsx do
  #       response.headers['Content-Disposition'] = 'attachment; filename=Отчет.xlsx'
  #       render xlsx: 'Докладная записка', template: 'over_work/do_dokladnaya'
  #     end
  #   end
  # end
  
  private

  def load_form_data
    @users = User.all
  end
end
